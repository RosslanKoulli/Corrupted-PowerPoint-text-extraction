#!/usr/bin/env python3
"""Reconstruct a corrupt .pptx by recovering its original parts.

I built this module around one idea about fidelity. A slide's layout lives
entirely inside its own original XML. For every shape and picture, the slide
records its position as a transform in English Metric Units, where 914400 units
equal one inch:

    <a:xfrm><a:off x="838200" y="365125"/><a:ext cx="7772400" cy="1325563"/></a:xfrm>

A picture's link to its image file lives in the slide's relationships:

    <Relationship Id="rId3" Type=".../image" Target="../media/image7.png"/>

Because the positions are stored as text inside the part, the only way to get
the original layout back is to recover those original parts verbatim rather
than regenerate them. Any approach that rebuilds slide XML from scratch invents
new coordinates and destroys the layout. I avoid that entirely.

My recovery proceeds in six steps. First I scan the raw bytes for every Local
File Header, whose signature is PK\\x03\\x04. Second, for each header I read the
stored filename so the original directory tree is preserved. Third, I inflate
the data when it is deflate-compressed or copy it when it is stored, which
gives me the original bytes of that part. Fourth, I write all recovered parts
into a directory tree. Fifth, I synthesize only the two mandatory plumbing
parts when they are missing, namely [Content_Types].xml and _rels/.rels, so the
package can open while everything carrying real layout is kept as is. Sixth, I
re-zip the parts into a clean .pptx with a fresh central directory and EOCD.

After writing the file I print validation commands so the result can be
confirmed with a strict loader and a renderer.

This module depends only on the standard library: struct, zlib, and zipfile.
No third-party packages are required for recovery itself.
"""

import sys
import os
import struct
import zlib
import zipfile
import argparse
import shutil
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger("pptx_recover")

SIG_LOCAL_FILE = b'PK\x03\x04'
SIG_DATA_DESC  = b'PK\x07\x08'

# Content types are FIXED by OOXML part type, so when [Content_Types].xml is
# lost I can rebuild it deterministically from the part names I recovered.
# These are the standard PresentationML overrides.
CONTENT_TYPE_OVERRIDES = {
    'ppt/presentation.xml':
        'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'ppt/slides/slide':
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    'ppt/slideLayouts/slideLayout':
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    'ppt/slideMasters/slideMaster':
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    'ppt/theme/theme':
        'application/vnd.openxmlformats-officedocument.theme+xml',
    'ppt/notesSlides/notesSlide':
        'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    'ppt/notesMasters/notesMaster':
        'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
}

# Default extension -> content type, for the <Default> entries.
DEFAULT_EXTENSIONS = {
    'rels': 'application/vnd.openxmlformats-package.relationships+xml',
    'xml':  'application/xml',
    'png':  'image/png',
    'jpeg': 'image/jpeg',
    'jpg':  'image/jpeg',
    'gif':  'image/gif',
    'bmp':  'image/bmp',
    'tiff': 'image/tiff',
    'tif':  'image/tiff',
    'emf':  'image/x-emf',
    'wmf':  'image/x-wmf',
    'svg':  'image/svg+xml',
    'mp4':  'video/mp4',
    'mov':  'video/quicktime',
    'mp3':  'audio/mpeg',
    'wav':  'audio/wav',
    'm4a':  'audio/mp4',
}


def parse_local_header(data, offset):
    """Parse a Local File Header at `offset` into a dict, or None if invalid.

    This is the recovery counterpart to the diagnostic parser. I read the same
    30-byte fixed header (signature, flags, compression method, sizes, and the
    filename length) and return the fields I need to extract the part: its
    name, its compression method, whether it uses a data descriptor, and the
    byte offset where its data begins.

    Args:
        data: The full archive bytes.
        offset: The byte offset where the header is expected to start.

    Returns:
        A dict with the parsed fields and a `data_start` offset, or None when
        the offset does not hold a valid header.
    """
    if offset + 30 > len(data) or data[offset:offset+4] != SIG_LOCAL_FILE:
        return None
    (version, flags, method, mtime, mdate,
     crc, comp_size, uncomp_size,
     name_len, extra_len) = struct.unpack('<HHHHHIIIHH', data[offset+4:offset+30])
    name_start = offset + 30
    name_end = name_start + name_len
    if name_end > len(data):
        return None
    name = data[name_start:name_end].decode('utf-8', errors='replace')
    return {
        'offset': offset,
        'flags': flags,
        'has_data_descriptor': bool(flags & 0x08),
        'method': method,
        'crc': crc,
        'comp_size': comp_size,
        'uncomp_size': uncomp_size,
        'name': name,
        'data_start': name_end + extra_len,
    }


def inflate_stored_or_deflate(data, header, next_header_offset):
    """Return the original decompressed bytes of a part, or None on failure.

    I handle the two compression methods that the OOXML packaging profile
    allows. When the method is 0 the bytes are stored uncompressed, so I copy
    them directly; PowerPoint stores already-compressed media such as PNG and
    JPEG this way. When the method is 8 the bytes are raw DEFLATE with no zlib
    header, so I decompress with wbits set to -15, which tells zlib there is no
    header to expect. XML parts such as slides, relationships, and themes are
    stored this way.

    Boundary handling matters because I need to know where each part's data
    ends. Normally the header records the real compressed size, so I slice
    exactly that many bytes. If the header used a data descriptor and deferred
    its size, the size field can be zero. In that case I decompress greedily
    from the data start up to the next header, relying on the fact that a
    DEFLATE stream is self-terminating, so the decompressor stops on its own at
    the true end of the stream.

    Args:
        data: The full archive bytes.
        header: A parsed header dict from parse_local_header.
        next_header_offset: The offset of the following header, or None for the
            last part. I use it as the outer bound when the size is unknown.

    Returns:
        The original bytes of the part, or None if the stream cannot be
        decompressed.
    """
    start = header['data_start']

    if header['method'] == 0:  # STORED
        if header['comp_size'] > 0:
            return data[start:start + header['comp_size']]
        # Stored + data descriptor (rare): copy up to the next header.
        end = next_header_offset if next_header_offset else len(data)
        return data[start:end]

    if header['method'] == 8:  # DEFLATE
        if header['comp_size'] > 0 and not header['has_data_descriptor']:
            blob = data[start:start + header['comp_size']]
            try:
                return zlib.decompress(blob, -15)  # -15 = raw deflate
            except zlib.error as e:
                log.warning(f"  deflate failed for {header['name']}: {e}")
                return None
        # No reliable size: stream-decompress until the deflate block ends.
        # The decompressor stops at the end of the stream and tells us how many
        # bytes it actually consumed, so I don't need the size up front.
        end = next_header_offset if next_header_offset else len(data)
        decomp = zlib.decompressobj(-15)
        try:
            out = decomp.decompress(data[start:end])
            out += decomp.flush()
            return out
        except zlib.error as e:
            log.warning(f"  streaming deflate failed for {header['name']}: {e}")
            return None

    log.warning(f"  unsupported compression method {header['method']} for {header['name']}")
    return None


def scan_and_recover(input_path):
    """Scan the whole file for local headers and recover every part possible.

    I first collect every local-header offset, because each part needs to know
    where the next header begins for the unknown-size fallback in
    inflate_stored_or_deflate. Then I parse each header, skip directory entries
    and unparseable headers, and decompress the data. I store results in a dict
    keyed by part name, so if a name legitimately repeats, the last complete
    copy wins.

    Args:
        input_path: Path to the corrupt file to recover from.

    Returns:
        A dict mapping each recovered part name to its original bytes.
    """
    log.info(f"Reading {input_path} ...")
    with open(input_path, 'rb') as f:
        data = f.read()
    log.info(f"  {len(data):,} bytes in memory")

    # Find every local-header offset first, so each part knows where the NEXT
    # one begins (needed for the data-descriptor / unknown-size fallback).
    offsets = []
    pos = 0
    while True:
        idx = data.find(SIG_LOCAL_FILE, pos)
        if idx == -1:
            break
        offsets.append(idx)
        pos = idx + 1
    log.info(f"  found {len(offsets)} local-header signatures")

    recovered = {}
    skipped = 0
    for i, off in enumerate(offsets):
        header = parse_local_header(data, off)
        if not header or not header['name'] or header['name'].endswith('/'):
            # Skip unparseable headers and directory entries (names ending in /).
            continue
        next_off = offsets[i + 1] if i + 1 < len(offsets) else None
        content = inflate_stored_or_deflate(data, header, next_off)
        if content is None:
            skipped += 1
            continue
        recovered[header['name']] = content

    log.info(f"  recovered {len(recovered)} parts ({skipped} streams unrecoverable)")
    return recovered


def synthesize_content_types(part_names):
    """Build a valid [Content_Types].xml from the recovered part names.

    I only call this when the original content-types part was missing or could
    not be recovered. I can rebuild it deterministically because the content
    type of an OOXML part is fixed by its type: I emit one Default entry per
    file extension present and one Override entry per typed part.

    Args:
        part_names: An iterable of the recovered part names.

    Returns:
        The content-types XML encoded as bytes.
    """
    # Collect the file extensions I actually have, for <Default> entries.
    exts = set()
    for name in part_names:
        ext = name.rsplit('.', 1)[-1].lower() if '.' in name else ''
        if ext:
            exts.add(ext)

    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']

    # Always include rels + xml defaults, then any media/other extensions I know.
    for ext in sorted(exts | {'rels', 'xml'}):
        ct = DEFAULT_EXTENSIONS.get(ext)
        if ct:
            lines.append(f'  <Default Extension="{ext}" ContentType="{ct}"/>')

    # Per-part overrides for the typed PresentationML parts.
    for name in sorted(part_names):
        part = '/' + name
        ct = None
        if name == 'ppt/presentation.xml':
            ct = CONTENT_TYPE_OVERRIDES['ppt/presentation.xml']
        else:
            for prefix, ctype in CONTENT_TYPE_OVERRIDES.items():
                if prefix.endswith(('slide', 'slideLayout', 'slideMaster',
                                    'theme', 'notesSlide', 'notesMaster')) \
                        and name.startswith(prefix) and name.endswith('.xml'):
                    ct = ctype
                    break
        if ct:
            lines.append(f'  <Override PartName="{part}" ContentType="{ct}"/>')

    lines.append('</Types>')
    return "\n".join(lines).encode('utf-8')


def synthesize_root_rels():
    """Build the package-level _rels/.rels, used only when it was lost.

    This root relationship part needs just one entry: it names
    ppt/presentation.xml as the start part of the presentation, which is how a
    reader knows where the document begins.

    Returns:
        The root relationships XML encoded as bytes.
    """
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
           '  <Relationship Id="rId1" '
           'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
           'Target="ppt/presentation.xml"/>\n'
           '</Relationships>')
    return xml.encode('utf-8')


def repair_opc_roots(recovered):
    """Ensure the two mandatory OPC root parts exist, preferring originals.

    The package cannot open without [Content_Types].xml and _rels/.rels. I
    synthesize each one only when it was not recovered, so a genuine original
    is always kept over a synthesized substitute. I mutate the `recovered` dict
    in place and return notes describing what I did.

    Args:
        recovered: The dict of recovered parts, modified in place.

    Returns:
        A list of short strings describing which roots were kept or synthesized.
    """
    notes = []
    if '[Content_Types].xml' not in recovered:
        recovered['[Content_Types].xml'] = synthesize_content_types(recovered.keys())
        notes.append("[Content_Types].xml was missing -> synthesized from part names")
    else:
        notes.append("[Content_Types].xml recovered from original (kept as-is)")

    if '_rels/.rels' not in recovered:
        recovered['_rels/.rels'] = synthesize_root_rels()
        notes.append("_rels/.rels was missing -> synthesized (points to ppt/presentation.xml)")
    else:
        notes.append("_rels/.rels recovered from original (kept as-is)")

    return notes


def summarize(recovered):
    """Print a count of recovered parts grouped by type, for a sanity check.

    I bucket each part into a category such as slides, layouts, masters, theme,
    media, or relationships, and print the totals. This lets the caller confirm
    at a glance that the expected parts came back.

    Args:
        recovered: The dict of recovered parts.
    """
    cats = {}
    for name in recovered:
        if name == '[Content_Types].xml':
            k = '[Content_Types].xml'
        elif name.startswith('ppt/slides/') and name.endswith('.xml') and '_rels' not in name:
            k = 'slides'
        elif name.startswith('ppt/slideLayouts/') and name.endswith('.xml'):
            k = 'slideLayouts'
        elif name.startswith('ppt/slideMasters/') and name.endswith('.xml'):
            k = 'slideMasters'
        elif name.startswith('ppt/theme/'):
            k = 'theme'
        elif name.startswith('ppt/media/'):
            k = 'media'
        elif name.endswith('.rels'):
            k = 'rels'
        elif name == 'ppt/presentation.xml':
            k = 'presentation.xml'
        else:
            k = 'other'
        cats[k] = cats.get(k, 0) + 1

    log.info("Recovered parts by type:")
    for k in ('[Content_Types].xml', 'presentation.xml', 'slides', 'slideLayouts',
              'slideMasters', 'theme', 'rels', 'media', 'other'):
        if k in cats:
            log.info(f"    {k:20s}: {cats[k]}")

    # A faithful reconstruction needs slides AND their rels AND layouts/masters.
    if not cats.get('slides'):
        log.warning("  No slide XML recovered; layout cannot be reconstructed from this file.")
    if cats.get('slides') and not cats.get('rels'):
        log.warning("  Slides recovered but no .rels; images may not bind to their positions.")


def write_tree(recovered, tree_dir):
    """Write every recovered part to disk under tree_dir, preserving paths.

    I recreate the original directory structure so the extracted tree can be
    inspected directly. I create parent directories as needed.

    Args:
        recovered: The dict of recovered parts.
        tree_dir: The directory to write the parts into.
    """
    for name, content in recovered.items():
        dest = os.path.join(tree_dir, name)
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        with open(dest, 'wb') as f:
            f.write(content)


def repackage(recovered, out_path):
    """Re-zip the recovered parts into a clean, structurally valid .pptx.

    Python's zipfile writes a fresh central directory and EOCD and recomputes
    every CRC, so the output is a valid ZIP no matter how broken the input was.
    I write [Content_Types].xml first as a matter of hygiene even though OPC
    readers tolerate any part order, and I store already-compressed media
    without recompressing it, which is both smaller and faster.

    Args:
        recovered: The dict of recovered parts.
        out_path: Where to write the rebuilt .pptx.
    """
    names = list(recovered.keys())
    names.sort(key=lambda n: (n != '[Content_Types].xml', n))  # content types first
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name in names:
            # Store already-compressed media without recompressing (smaller + faster).
            ext = name.rsplit('.', 1)[-1].lower() if '.' in name else ''
            compress = (zipfile.ZIP_STORED
                        if ext in ('png', 'jpg', 'jpeg', 'gif', 'mp4', 'mov', 'mp3')
                        else zipfile.ZIP_DEFLATED)
            z.writestr(name, recovered[name], compress_type=compress)
    log.info(f"Wrote {out_path} ({os.path.getsize(out_path):,} bytes)")


def main():
    """Parse arguments, run the chosen mode, and return a process exit code.

    In validate mode I only try to open the input and report whether it loads
    and whether its CRCs pass, changing nothing. In scan mode I run the full
    recovery: scan for parts, repair the OPC roots, optionally write the
    extracted tree, and repackage into the output file. I return 0 on success
    and a non-zero code when the input is missing or nothing could be recovered.
    """
    parser = argparse.ArgumentParser(
        description="Faithfully reconstruct a corrupt .pptx by recovering its "
                    "original parts from surviving ZIP local headers.")
    parser.add_argument('input', help="Path to the corrupt .pptx file")
    parser.add_argument('--out', '-o', default='fixed.pptx',
                        help="Output path for the rebuilt .pptx (default: fixed.pptx)")
    parser.add_argument('--mode', choices=['scan', 'validate'], default='scan',
                        help="'scan' = recover from local headers (default); "
                             "'validate' = just try to open INPUT and report (no rebuild)")
    parser.add_argument('--keep-tree', default=None,
                        help="Also write the recovered parts to this directory for inspection")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        log.error(f"file not found: {args.input}")
        return 2

    if args.mode == 'validate':
        # Just check whether the file opens and whether CRCs pass.
        try:
            with zipfile.ZipFile(args.input) as zf:
                names = zf.namelist()
                bad = zf.testzip()
            log.info(f"Opened OK. {len(names)} parts. "
                     f"{'all CRCs OK' if bad is None else 'first bad part: ' + bad}")
            return 0
        except Exception as e:
            log.error(f"Cannot open as ZIP: {e}")
            log.error("Use --mode scan to attempt structural recovery.")
            return 1

    # ---- scan/recover pipeline
    recovered = scan_and_recover(args.input)
    if not recovered:
        log.error("Nothing recoverable. The file may be encrypted, not a ZIP, or "
                  "severely damaged. Prioritize finding the original copy.")
        return 1

    notes = repair_opc_roots(recovered)
    summarize(recovered)
    for n in notes:
        log.info(f"OPC roots: {n}")

    if args.keep_tree:
        os.makedirs(args.keep_tree, exist_ok=True)
        write_tree(recovered, args.keep_tree)
        log.info(f"Recovered tree written to {args.keep_tree}/")

    repackage(recovered, args.out)

    # ---- next-step validation guidance
    print("\n" + "=" * 72)
    print("RECOVERY COMPLETE. Now validate fidelity (run these):")
    print("=" * 72)
    print(f"  1) Load test:")
    print(f'       python -c "from pptx import Presentation; '
          f"Presentation('{args.out}'); print('python-pptx OK')\"")
    print(f"  2) Render test (positions + images):")
    print(f"       soffice --headless --convert-to pdf --outdir . '{args.out}'")
    print(f"     then open the PDF and compare layout to what you remember.")
    print(f"  3) If a slide is missing or blank, inspect the recovered tree:")
    print(f"       (re-run with --keep-tree recovered_parts/ and look in")
    print(f"        recovered_parts/ppt/slides/ and .../slides/_rels/)")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
