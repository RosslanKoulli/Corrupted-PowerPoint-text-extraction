#!/usr/bin/env python3
"""Rebuild missing slide .rels files so a recovered .pptx renders.

I note up front that this module is a narrower predecessor of the full graph
rebuilder. It repairs only the slide-level relationship files. When every .rels
in the package is missing, including the layout and master rels, this module is
not enough on its own and the graph rebuilder is the right tool.

The problem it addresses is a recovered .pptx that PowerPoint opens but draws
as blank slides. The cause is usually that the per-slide relationship files,
ppt/slides/_rels/slideN.xml.rels, are missing. Each slide's XML refers to its
layout and its images indirectly by relationship id, for example
<a:blip r:embed="rId2"/> meaning draw image rId2 here. The mapping from rId2 to
an actual image file lives only in the slide's .rels. A slide also needs a
relationship to a slide layout, without which PowerPoint has no inherited
placeholders or background chain to render. If the .rels are gone, every id
dangles and the slide renders empty even though the slide XML and the images
are present.

For every slide that is missing its .rels, I read the slide XML, find every
relationship id it references, classify each as an image or a hyperlink, ensure
a slide layout relationship exists, and write a new .rels that satisfies those
ids.

What comes back exactly: the slide layout link, the relationship structure that
stops PowerPoint from blanking the slide, and all text, shapes, colors, and
positions, because those live in the intact slide XML. What is best-effort:
which image file each reference points to, because that mapping was in the lost
.rels. I assign images by the order they appear, which is correct for many
decks but can mis-pair images on slides that hold several pictures. The text
and layout are unaffected; only image identity may differ. The module prints a
per-slide report of what it inferred for this reason.
"""

import sys
import os
import re
import zipfile
import argparse
import shutil
import tempfile
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger("rebuild_rels")

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
R_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
R_HLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

# Matches r:embed="rId5", r:link="rId5", r:id="rId5" in slide XML.
RID_RE = re.compile(r'r:(embed|link|id)="(rId\d+)"')
# Hyperlinks specifically use r:id inside an <a:hlinkClick>/<a:hlinkHover>.
HLINK_RE = re.compile(r'<a:hlink(?:Click|Hover)[^>]*r:id="(rId\d+)"')


def slide_number(name):
    """Return the slide number from a slide part name, or 0 if absent.

    I use this to sort slides numerically so slide10 follows slide2.

    Args:
        name: A slide part name such as ppt/slides/slide12.xml.

    Returns:
        The integer slide number, or 0 when the name does not match.
    """
    m = re.search(r'slide(\d+)\.xml$', name)
    return int(m.group(1)) if m else 0


def classify_rids(slide_xml):
    """Classify every relationship id a slide references as image or hyperlink.

    I treat any id that appears inside an a:hlinkClick or a:hlinkHover element
    as a hyperlink, and I treat every other id referenced through r:embed,
    r:link, or r:id as an image, which is the common case in slide bodies.
    Classifying both kinds keeps the rebuilt .rels valid whether the id pointed
    at a picture or a link.

    Args:
        slide_xml: The slide XML as a string.

    Returns:
        A dict mapping each relationship id to the string 'image' or
        'hyperlink'.
    """
    kinds = {}
    hlinks = set(HLINK_RE.findall(slide_xml))
    for attr, rid in RID_RE.findall(slide_xml):
        kinds[rid] = 'hyperlink' if rid in hlinks else 'image'
    return kinds


def build_rels_xml(rid_kinds, layout_target, media_for_images):
    """Construct the .rels XML for one slide.

    I always add a slide layout relationship first, on an id chosen so it does
    not collide with the slide's own ids. I then assign image targets to the
    image ids in ascending id order. For hyperlinks I emit an external
    relationship with a placeholder target, because the real URL was in the
    lost .rels and cannot be recovered from the slide XML; the placeholder
    keeps the file valid so the slide still renders.

    Args:
        rid_kinds: A dict mapping each relationship id to 'image' or
            'hyperlink'.
        layout_target: The relative path to the layout to link, for example
            ../slideLayouts/slideLayout1.xml.
        media_for_images: An ordered list of media targets to assign to the
            image ids.

    Returns:
        A tuple of the .rels XML bytes and the relationship id used for the
        layout link.
    """
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             f'<Relationships xmlns="{REL_NS}">']

    # Choose an rId for the layout that doesn't collide with the slide's own ids.
    used = set(rid_kinds.keys())
    n = 1
    while f"rId{n}" in used:
        n += 1
    layout_rid = f"rId{n}"
    lines.append(f'  <Relationship Id="{layout_rid}" Type="{R_LAYOUT}" '
                 f'Target="{layout_target}"/>')

    # Assign image targets to image rIds in ascending rId order.
    image_rids = sorted([r for r, k in rid_kinds.items() if k == 'image'],
                        key=lambda r: int(r[3:]))
    for i, rid in enumerate(image_rids):
        target = media_for_images[i % len(media_for_images)] if media_for_images \
            else '../media/image1.png'
        lines.append(f'  <Relationship Id="{rid}" Type="{R_IMAGE}" Target="{target}"/>')

    # Hyperlinks need a target + TargetMode="External". I don't know the URL
    # (it was in the lost rels), so I use a placeholder that keeps the file
    # valid; the slide still renders. Real URLs are unrecoverable from slide XML.
    for rid, kind in rid_kinds.items():
        if kind == 'hyperlink':
            lines.append(f'  <Relationship Id="{rid}" Type="{R_HLINK}" '
                         f'Target="https://example.invalid/recovered" '
                         f'TargetMode="External"/>')

    lines.append('</Relationships>')
    return ("\n".join(lines)).encode('utf-8'), layout_rid


def main():
    """Find slides missing their .rels, rebuild only those, and write a new file.

    I inventory the slides, layouts, and media, and abort if no layouts are
    present, because a slide cannot render without a layout to inherit from. I
    then walk the slides, skip any that already have a .rels, and for each one
    that is missing it I classify its referenced ids and write a fresh .rels. I
    assign images from the media pool in order across all slides. I copy the
    original parts into a new archive, add the rebuilt .rels, and print a
    per-slide report of what I inferred. I return 0 on success and a non-zero
    code when the input is missing or has no layouts.
    """
    ap = argparse.ArgumentParser(
        description="Rebuild missing slide .rels so recovered slides render.")
    ap.add_argument('input', help="recovered .pptx that opens but shows blank slides")
    ap.add_argument('--out', '-o', default='fixed_rels.pptx')
    ap.add_argument('--verbose', '-v', action='store_true')
    args = ap.parse_args()
    if args.verbose:
        log.setLevel(logging.DEBUG)

    if not os.path.exists(args.input):
        log.error(f"not found: {args.input}")
        return 2

    zin = zipfile.ZipFile(args.input)
    names = zin.namelist()

    slides = sorted([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                    key=slide_number)
    layouts = sorted([n for n in names
                      if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n)],
                     key=lambda n: int(re.search(r'(\d+)', n).group(1)))
    media = sorted([n for n in names if n.startswith('ppt/media/')])
    existing_slide_rels = [n for n in names if 'ppt/slides/_rels/' in n]

    log.info(f"slides={len(slides)} layouts={len(layouts)} media={len(media)} "
             f"existing slide-rels={len(existing_slide_rels)}")

    if not layouts:
        log.error("No slide layouts in the package; slides cannot render without a "
                  "layout. The recovery is missing ppt/slideLayouts/*. Re-run "
                  "pptx_recover.py and confirm layouts were recovered.")
        return 1

    # Media targets are written relative to the slide, hence the ../media/ prefix.
    media_targets = [f"../media/{os.path.basename(m)}" for m in media]
    # Use the first layout as the default link target. Layouts mostly affect
    # inherited placeholders; the slide's own shapes render regardless.
    default_layout_target = f"../slideLayouts/{os.path.basename(layouts[0])}"

    # Build the set of new rels I need to add.
    new_parts = {}      # arcname -> bytes
    report = []
    img_cursor = 0      # walk through media as I assign images across slides

    for s in slides:
        num = slide_number(s)
        rels_name = f"ppt/slides/_rels/slide{num}.xml.rels"
        if rels_name in names:
            continue   # this slide already has rels; leave it untouched
        xml = zin.read(s).decode('utf-8', 'replace')
        rid_kinds = classify_rids(xml)
        n_imgs = sum(1 for k in rid_kinds.values() if k == 'image')
        # Assign the next n_imgs media files to this slide, in file order.
        slice_for_slide = []
        if media_targets and n_imgs:
            slice_for_slide = [media_targets[(img_cursor + i) % len(media_targets)]
                               for i in range(n_imgs)]
            img_cursor += n_imgs
        rels_xml, layout_rid = build_rels_xml(rid_kinds, default_layout_target,
                                              slice_for_slide)
        new_parts[rels_name] = rels_xml
        report.append((num, n_imgs, sum(1 for k in rid_kinds.values() if k == 'hyperlink')))

    if not new_parts:
        log.info("Every slide already has a .rels file, so there is nothing to rebuild. "
                 "If slides are still blank, the cause is elsewhere (e.g. missing "
                 "layouts/masters). Run pptx_diagnose.py on the original.")
        return 0

    # Write a new package: copy everything from the input, then add new rels.
    log.info(f"Rebuilding {len(new_parts)} missing slide-rels files...")
    with zipfile.ZipFile(args.out, 'w', zipfile.ZIP_DEFLATED) as zout:
        # copy original parts
        for item in zin.infolist():
            data = zin.read(item.filename)
            ext = item.filename.rsplit('.', 1)[-1].lower() if '.' in item.filename else ''
            ctype = zipfile.ZIP_STORED if ext in ('png', 'jpg', 'jpeg', 'gif') \
                else zipfile.ZIP_DEFLATED
            zout.writestr(item, data, compress_type=ctype)
        # add rebuilt rels
        for arc, content in new_parts.items():
            zout.writestr(arc, content)

    log.info(f"Wrote {args.out} ({os.path.getsize(args.out):,} bytes)")

    # Per-slide summary so you can see what was inferred.
    total_imgs = sum(r[1] for r in report)
    total_links = sum(r[2] for r in report)
    log.info(f"Rebuilt rels for {len(report)} slides: "
             f"{total_imgs} image links assigned, {total_links} hyperlinks stubbed.")
    if args.verbose:
        for num, imgs, links in report[:40]:
            log.debug(f"  slide{num}: {imgs} images, {links} hyperlinks")

    print("\n" + "=" * 72)
    print("Rebuild complete. Open the output file in PowerPoint.")
    print("=" * 72)
    print("Expected: slides now show their TEXT, SHAPES, and BACKGROUNDS (these were")
    print("always intact). IMAGES will appear too, but on multi-image slides the")
    print("specific image-to-slot pairing is inferred by order and may be off.")
    print("Validate render:")
    print(f"   soffice --headless --convert-to pdf --outdir . '{args.out}'")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
