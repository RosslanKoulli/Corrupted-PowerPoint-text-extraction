#!/usr/bin/env python3
"""Diagnose the structural damage in a corrupt .pptx file (read-only).

I wrote this module because a .pptx is a ZIP archive, and ZIP readers
(PowerPoint, unzip, Python's zipfile) never read the file front to back. They
jump to the very end of the file and read the End Of Central Directory record,
whose signature is the four bytes PK\\x05\\x06. That record points them at the
central directory, which is the index that lists every part inside the archive
and the byte offset where each one sits.

That design decides how I have to think about recovery. If the central
directory or the EOCD is damaged or missing, a reader cannot locate any part at
all, which is why PowerPoint reports that the file cannot be repaired. The part
data itself, meaning the slides, the images, and the layout, is usually still
intact in the body of the file. The index that points to it is the casualty,
not the content.

I do not repair anything in this module. I only inspect the bytes and decide
which of three damage situations applies, because each one calls for a
different recovery path:

    (a) Bad central directory, intact body -> almost everything is recoverable
    (b) Truncation, the tail is missing    -> the surviving prefix is recoverable
    (c) Isolated data corruption           -> directory is fine, a few parts are bad

The intended workflow is to run this module first, read the verdict it prints,
and then run the recovery module accordingly. This module prints a report to
stdout, never modifies the input, and exits 0.
"""

import sys
import os
import struct
import argparse
import zipfile
import time


#
# Progress bar (pure stdlib, no third-party deps).
#
# Why stderr and not stdout: the diagnosis REPORT goes to stdout, so if you
# run `python pptx_diagnose.py big.pptx > report.txt` the report stays clean
# and the animated bar still shows live in your terminal (stderr is not
# redirected). Why no tqdm: this script is a triage tool you may run on a
# bare machine; zero install requirements means it always works.
#
# The bar only draws if stderr is an interactive terminal (isatty). When the
# output is piped/redirected it stays silent so logs don't fill with control
# characters.
#
class ProgressBar:
    """A terminal progress bar that uses only the standard library.

    I draw to stderr rather than stdout on purpose. The diagnosis report goes
    to stdout, so writing the bar to stderr means a redirect like
    `python pptx_diagnose.py big.pptx > report.txt` keeps the report clean while
    the animated bar still shows live in the terminal. I avoid tqdm because this
    is a triage tool that may run on a bare machine, and depending on nothing
    means it always works.

    The bar only animates when stderr is an interactive terminal, which I test
    with isatty(). When the output is piped or redirected I stay silent so logs
    do not fill up with cursor-control characters.
    """

    def __init__(self, total, label, width=34, min_interval=0.1):
        """Set up a bar for `total` units of work labelled with `label`.

        I clamp `total` to at least 1 so the fraction math can never divide by
        zero on empty input. `min_interval` is the minimum seconds between
        redraws, which throttles drawing so I do not spend more time animating
        than working. I record the start time here so I can show elapsed
        seconds, and I draw the initial empty frame only if the bar is enabled.
        """
        self.total = max(1, total)       # avoid divide-by-zero on empty input
        self.label = label
        self.width = width
        self.min_interval = min_interval  # seconds between redraws (throttle)
        self.enabled = sys.stderr.isatty()
        self._last_draw = 0.0
        self._start = time.time()
        if self.enabled:
            self._draw(0)

    def update(self, current):
        """Redraw the bar at `current` out of total, subject to throttling.

        I return immediately when the bar is disabled. I always let the final
        frame through when `current` reaches total, so the bar visibly
        completes; otherwise I skip the redraw if less than `min_interval`
        seconds have passed since the last one.
        """
        if not self.enabled:
            return
        now = time.time()
        # Always allow the final 100% frame through; otherwise throttle.
        if current < self.total and (now - self._last_draw) < self.min_interval:
            return
        self._last_draw = now
        self._draw(current)

    def _draw(self, current):
        """Render one frame of the bar in place.

        I compute the filled fraction, build the bar from filled and empty
        block characters, and write it after a carriage return so each frame
        overwrites the previous one on the same line instead of scrolling.
        """
        filled = int(self.width * frac)
        bar = '█' * filled + '░' * (self.width - filled)
        elapsed = time.time() - self._start
        # \r returns the cursor to line start so I overwrite in place.
        sys.stderr.write(f"\r  {self.label:<22} |{bar}| {frac*100:5.1f}%  {elapsed:4.1f}s")
        sys.stderr.flush()

    def finish(self):
        """Snap the bar to 100 percent and end the line.

        I draw the full frame and then write a newline so that whatever prints
        next starts on a clean line instead of on top of the bar.
        """
        if not self.enabled:
            return
        self._draw(self.total)
        sys.stderr.write("\n")
        sys.stderr.flush()

#
# ZIP record signatures. These are 4-byte "magic numbers" that mark the start
# of each kind of ZIP record. I search for these as raw bytes in the file.
# The bytes are little-endian, which is why e.g. EOCD (0x06054b50) is written
# as the byte sequence 'P' 'K' 0x05 0x06.
#
SIG_LOCAL_FILE   = b'PK\x03\x04'   # Local File Header, precedes each part's data
SIG_CENTRAL_DIR  = b'PK\x01\x02'   # Central Directory File Header, one per part, in the index
SIG_EOCD         = b'PK\x05\x06'   # End Of Central Directory, the "table of contents" pointer
SIG_ZIP64_EOCD   = b'PK\x06\x06'   # Zip64 End Of Central Directory (for >4GB / >65535 entries)
SIG_ZIP64_LOC    = b'PK\x06\x07'   # Zip64 EOCD Locator
SIG_DATA_DESC    = b'PK\x07\x08'   # Optional Data Descriptor (sizes/CRC written AFTER the data)


def human_size(n):
    """Return a byte count formatted as a readable string.

    I step up through B, KB, MB, GB, and TB, dividing by 1024 each time, and
    stop at the first unit where the value is below 1024.
    """
    for unit in ('B', 'KB', 'MB', 'GB', 'TB'):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} PB"


def find_all(data, sig, limit=None, progress_label=None):
    """Find every byte offset where a 4-byte signature occurs in the data.

    I search in 16 MB chunks so I can update a progress bar on a large file
    instead of blocking silently. A signature can straddle a chunk boundary, so
    I extend each chunk's search window by len(sig)-1 bytes. To avoid reporting
    a boundary match twice, I stop counting once a match falls into that overlap
    region, because the next chunk will report it.

    Args:
        data: The raw bytes to search.
        sig: The signature bytes to look for.
        limit: Optional cap on how many offsets to collect before returning
            early, useful when I only need a rough count.
        progress_label: Optional label; when set, I show a progress bar.

    Returns:
        A list of integer byte offsets, in ascending order.
    """
    offsets = []
    n = len(data)
    overlap = len(sig) - 1
    # Chunk size: 16 MB balances redraw frequency against per-chunk overhead.
    chunk = 16 * 1024 * 1024
    bar = ProgressBar(n, progress_label) if progress_label else None

    start = 0
    while start < n:
        end = min(n, start + chunk)
        # Search this window (extended by `overlap` so boundary hits aren't lost).
        window_end = min(n, end + overlap)
        local = start
        while True:
            idx = data.find(sig, local, window_end)
            if idx == -1 or idx >= end:
                # Stop once matches fall into the overlap region; the next
                # chunk (which starts at `end`) will report them, avoiding dupes.
                break
            offsets.append(idx)
            local = idx + 1
            if limit and len(offsets) >= limit:
                if bar:
                    bar.finish()
                return offsets
        if bar:
            bar.update(end)
        start = end

    if bar:
        bar.finish()
    return offsets


def parse_local_header(data, offset):
    """Parse a Local File Header at `offset` into a dict, or None if invalid.

    A Local File Header is the per-part record that sits in front of each
    part's data. It matters for recovery because it contains everything I need
    to extract that one part on its own, which is what lets me recover parts
    even when the central directory is gone. I return None when the bytes at
    `offset` are not a valid header, so callers can skip false matches.

    The fixed portion is 30 bytes laid out like this:

        offset  size  field
        0       4     signature (PK\\x03\\x04)
        4       2     version needed to extract
        6       2     general purpose bit flag   <-- bit 3 (0x08) = "data descriptor"
        8       2     compression method         <-- 0 = stored, 8 = deflate
        10      2     last mod time
        12      2     last mod date
        14      4     CRC-32
        18      4     compressed size
        22      4     uncompressed size
        26      2     file name length (n)
        28      2     extra field length (m)
        30      n     file name
        30+n    m     extra field

    Args:
        data: The full archive bytes.
        offset: The byte offset where the header is expected to start.

    Returns:
        A dict of the parsed fields plus a `has_data_descriptor` convenience
        flag, or None if the offset does not hold a valid header.
    """
    if offset + 30 > len(data):
        return None
    if data[offset:offset+4] != SIG_LOCAL_FILE:
        return None

    # '<' = little-endian. H = uint16, I = uint32. This unpacks bytes 4..30.
    (version, flags, method, mtime, mdate,
     crc, comp_size, uncomp_size,
     name_len, extra_len) = struct.unpack('<HHHHHIIIHH', data[offset+4:offset+30])

    name_start = offset + 30
    name_end = name_start + name_len
    if name_end > len(data):
        return None
    try:
        name = data[name_start:name_end].decode('utf-8', errors='replace')
    except Exception:
        name = "<undecodable>"

    return {
        'offset': offset,
        'version': version,
        'flags': flags,
        'has_data_descriptor': bool(flags & 0x08),  # bit 3: sizes are AFTER the data
        'method': method,                            # 0 stored, 8 deflate
        'crc': crc,
        'comp_size': comp_size,
        'uncomp_size': uncomp_size,
        'name': name,
        'name_len': name_len,
        'extra_len': extra_len,
        'data_start': name_end + extra_len,
    }


def try_python_open(path):
    """Ask Python's strict zipfile to open the file and report what happens.

    This is the single most informative check I run. If the open succeeds, the
    central directory is intact and the container is not the problem, so I look
    inside the parts instead. If it raises BadZipFile, the directory or EOCD is
    the casualty. When it does open, I also call testzip(), which returns the
    name of the first part with a bad CRC, letting me separate isolated data
    corruption from structural damage.

    Args:
        path: Path to the file to test.

    Returns:
        A tuple (ok, message, namelist). `ok` is True when the file opened;
        `message` summarises the CRC result or the error; `namelist` is the
        list of part names when it opened, otherwise None.
    """
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            # testzip() returns the name of the first file with a bad CRC, or None.
            bad = zf.testzip()
            return True, ("all CRCs OK" if bad is None
                          else f"first bad-CRC part: {bad}"), names
    except zipfile.BadZipFile as e:
        return False, f"BadZipFile: {e}", None
    except Exception as e:
        return False, f"{type(e).__name__}: {e}", None


def classify(path):
    """Run the full diagnosis, print a report, and return a verdict string.

    I gather the four checks (strict open, EOCD presence, local-header scan,
    and part inventory) and combine them into one of the verdict codes
    NO_ZIP_DAMAGE, ISOLATED_CRC, TRUNCATION, BAD_DIRECTORY, or SEVERE. The
    verdict tells the caller which recovery path applies.

    Args:
        path: Path to the file to diagnose.

    Returns:
        A short verdict string naming the damage situation.
    """
    size = os.path.getsize(path)
    print("=" * 72)
    print(f"DIAGNOSING: {path}")
    print(f"File size : {size:,} bytes ({human_size(size)})")
    print("=" * 72)

    with open(path, 'rb') as f:
        # Read in chunks so I can show progress on large files. For an 800 MB
        # file a single f.read() can take a few seconds with no feedback.
        rbar = ProgressBar(size, "reading file")
        chunks = []
        read_so_far = 0
        while True:
            block = f.read(16 * 1024 * 1024)
            if not block:
                break
            chunks.append(block)
            read_so_far += len(block)
            rbar.update(read_so_far)
        rbar.finish()
        data = b''.join(chunks)

    # --- Step 1: does a strict reader open it?
    print("\n[1] Strict open test (Python zipfile)")
    ok, msg, names = try_python_open(path)
    if ok:
        print(f"    OPENS CLEANLY. {msg}")
        print(f"    Parts listed in central directory: {len(names)}")
        if 'all CRCs OK' not in msg:
            print("    -> The directory is fine but at least one part has a bad CRC.")
            print("    -> This is situation (c): ISOLATED DATA CORRUPTION.")
        else:
            print("    -> No structural damage detected by Python. If PowerPoint still")
            print("       refuses it, the problem is likely an INVALID OOXML PART, not")
            print("       the ZIP container. (Different fix, see notes at the end.)")
    else:
        print(f"    FAILS: {msg}")
        print("    -> The central directory or EOCD is unreadable. Continuing...")

    # --- Step 2: is there an EOCD near the end?
    print("\n[2] End Of Central Directory (EOCD) check")
    # The EOCD lives at the very end, but a "ZIP comment" of up to 65535 bytes
    # can follow it, so I scan the last 64KB + 22 bytes to be safe.
    tail_window = data[-(65535 + 22):] if size > (65535 + 22) else data
    eocd_in_tail = find_all(tail_window, SIG_EOCD)
    zip64_eocd = find_all(data, SIG_ZIP64_EOCD, limit=1)

    if eocd_in_tail:
        # Offset of the last EOCD within the whole file
        eocd_off = (size - len(tail_window)) + eocd_in_tail[-1]
        print(f"    EOCD signature FOUND near end at offset {eocd_off:,}.")
        # Parse the EOCD's pointer to the central directory.
        if eocd_off + 22 <= size:
            (disk, cd_disk, ent_disk, ent_total,
             cd_size, cd_offset, comment_len) = struct.unpack(
                 '<HHHHIIH', data[eocd_off+4:eocd_off+22])
            print(f"      entries (per EOCD)      : {ent_total}")
            print(f"      central dir size        : {cd_size:,} bytes")
            print(f"      central dir offset      : {cd_offset:,}")
            # Sanity-check the pointer: does it land on a real central-dir header?
            if cd_offset == 0xFFFFFFFF or cd_size == 0xFFFFFFFF:
                print("      -> Values are 0xFFFFFFFF sentinels: this is a ZIP64 archive.")
            elif cd_offset < size and data[cd_offset:cd_offset+4] == SIG_CENTRAL_DIR:
                print("      -> Pointer lands on a valid central-directory header. The")
                print("         EOCD itself looks consistent.")
            else:
                print("      -> Pointer does NOT land on a central-directory header.")
                print("         The EOCD is present but its offset is wrong (file may")
                print("         have had bytes added/removed, or the directory is damaged).")
    else:
        print("    EOCD signature NOT found near the end of the file.")
        print("    -> The tail of the file is missing or destroyed.")
        print("    -> Strong indicator of situation (b): TRUNCATION.")

    if zip64_eocd:
        print("    NOTE: a ZIP64 EOCD record is present, so tools used for recovery")
        print("          must be ZIP64-aware (Info-ZIP zip/unzip, 7-Zip, Python all are;")
        print("          macOS Archive Utility is NOT).")

    # --- Step 3: how many local file headers survive in the body?
    print("\n[3] Local File Header scan (the actual recoverable parts)")
    local_offsets = find_all(data, SIG_LOCAL_FILE, progress_label="scanning headers")
    print(f"    Local File Header signatures found: {len(local_offsets)}")
    # Parse them and tally what's valid + categorize by path.
    parsed = []
    data_descriptor_count = 0
    methods = {}
    categories = {
        'slides': 0, 'slideLayouts': 0, 'slideMasters': 0,
        'theme': 0, 'media': 0, 'rels': 0, 'content_types': 0,
        'presentation': 0, 'other': 0,
    }
    pbar = ProgressBar(len(local_offsets), "parsing headers") if local_offsets else None
    for li, off in enumerate(local_offsets):
        if pbar:
            pbar.update(li)
        h = parse_local_header(data, off)
        if not h:
            continue
        parsed.append(h)
        if h['has_data_descriptor']:
            data_descriptor_count += 1
        methods[h['method']] = methods.get(h['method'], 0) + 1
        n = h['name']
        if n == '[Content_Types].xml':
            categories['content_types'] += 1
        elif n.startswith('ppt/slides/') and n.endswith('.xml') and '_rels' not in n:
            categories['slides'] += 1
        elif n.startswith('ppt/slideLayouts/'):
            categories['slideLayouts'] += 1
        elif n.startswith('ppt/slideMasters/'):
            categories['slideMasters'] += 1
        elif n.startswith('ppt/theme/'):
            categories['theme'] += 1
        elif n.startswith('ppt/media/'):
            categories['media'] += 1
        elif n.endswith('.rels'):
            categories['rels'] += 1
        elif n == 'ppt/presentation.xml':
            categories['presentation'] += 1
        else:
            categories['other'] += 1
    if pbar:
        pbar.finish()

    print(f"    Valid (parseable) local headers   : {len(parsed)}")
    if methods:
        method_names = {0: 'stored', 8: 'deflate'}
        pretty = ", ".join(f"{method_names.get(m, m)}={c}" for m, c in sorted(methods.items()))
        print(f"    Compression methods               : {pretty}")
    print(f"    Headers using data-descriptor mode: {data_descriptor_count}")
    if data_descriptor_count == 0 and parsed:
        print("      -> GOOD: every header carries its real compressed size inline.")
        print("         The recovery scanner can trust those sizes directly (typical")
        print("         for files saved by PowerPoint desktop / Office).")
    elif data_descriptor_count:
        print("      -> Some headers defer their sizes to a trailing data descriptor.")
        print("         The recovery scanner handles this, but boundaries are inferred.")

    print("\n    Recoverable OOXML parts by type (from surviving local headers):")
    for k in ('content_types', 'presentation', 'slides', 'slideLayouts',
              'slideMasters', 'theme', 'rels', 'media', 'other'):
        print(f"        {k:16s}: {categories[k]}")

    # Report how many media parts are still addressable via local headers.
    if categories['media']:
        print(f"\n    ({categories['media']} media parts are still addressable via")
        print(f"     local headers in the body of the file.)")

    # --- Step 4: verdict
    print("\n" + "=" * 72)
    print("VERDICT")
    print("=" * 72)

    verdict = "unknown"
    have_roots = categories['content_types'] and categories['presentation']

    if ok and 'all CRCs OK' in msg:
        verdict = "NO_ZIP_DAMAGE"
        print("The ZIP container is structurally fine and all CRCs pass.")
        print("If PowerPoint still won't open it, the problem is inside an OOXML")
        print("part (malformed XML), not the container. Use the OOXML-validation")
        print("path in pptx_recover.py (--mode validate) rather than ZIP repair.")
    elif ok and 'bad-CRC' in msg:
        verdict = "ISOLATED_CRC"
        print("Situation (c): the directory is readable but specific part(s) have")
        print("bad CRCs. Extract everything, then replace only the damaged part(s).")
        print("Run:  python pptx_recover.py INPUT --mode scan --out fixed.pptx")
    elif not eocd_in_tail and len(parsed) > 0:
        verdict = "TRUNCATION"
        print("Situation (b): TRUNCATION. The EOCD/central directory at the tail is")
        print("gone, but local headers survive in the body. Everything written")
        print("BEFORE the cut is recoverable; slides after the cut are lost.")
        print(f"Root parts present: {'YES' if have_roots else 'NO (will be synthesized)'}.")
        print("Run:  python pptx_recover.py INPUT --mode scan --out fixed.pptx")
    elif not ok and len(parsed) > 0:
        verdict = "BAD_DIRECTORY"
        print("Situation (a): BAD CENTRAL DIRECTORY but the body is intact.")
        print("This is the most recoverable case, and nearly everything should come back.")
        print(f"Root parts present: {'YES' if have_roots else 'NO (will be synthesized)'}.")
        print("Try the cheap fix first:")
        print('    zip -FF INPUT --out fixed.zip   (then rename fixed.zip -> fixed.pptx)')
        print("If that is incomplete, run the scanner:")
        print("    python pptx_recover.py INPUT --mode scan --out fixed.pptx")
    else:
        verdict = "SEVERE"
        print("Could not find usable local headers. The file may be severely")
        print("damaged, encrypted, or not actually a ZIP/PPTX. Recovery prospects")
        print("are poor; prioritize locating an original copy of the file from")
        print("any available backup or version history.")

    print("\nVerdict code:", verdict)
    print("=" * 72)
    return verdict


def main():
    """Parse arguments, run the diagnosis, and return a process exit code.

    I return 2 when the input file does not exist and 0 once the report has
    printed. I do not propagate the verdict into the exit code, because the
    verdict is meant to be read from the report rather than branched on by a
    shell.
    """
    parser = argparse.ArgumentParser(
        description="Diagnose the damage type in a corrupt .pptx (ZIP) file. "
                    "Read-only; nothing is modified.")
    parser.add_argument('input', help="Path to the corrupt .pptx file")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"ERROR: file not found: {args.input}", file=sys.stderr)
        return 2

    classify(args.input)
    return 0


if __name__ == "__main__":
    sys.exit(main())
