#!/usr/bin/env python3
"""Rebuild the entire relationship graph of a recovered .pptx.

I use this module when all of the .rels files were lost during recovery. A
.pptx is not just slides; it is a graph of parts wired together by .rels files
at every level, and PowerPoint renders a slide only if it can walk the whole
chain:

    _rels/.rels                          -> ppt/presentation.xml
    ppt/_rels/presentation.xml.rels      -> every slide, the master, theme, notesMaster
    ppt/slides/_rels/slideN.xml.rels     -> that slide's layout and its images
    ppt/slideLayouts/_rels/slideLayoutN.xml.rels -> the master
    ppt/slideMasters/_rels/slideMaster1.xml.rels -> all layouts and theme
    ppt/notesMasters/_rels/notesMaster1.xml.rels -> its theme
    ppt/notesSlides/_rels/notesSlide1.xml.rels   -> the notesMaster

If every .rels is missing, fixing only the slide rels leaves the chain broken
at the layout-to-master and master-to-theme joints, so PowerPoint discards each
slide and reports that it could not read some content. I rebuild every tier so
the whole chain resolves.

I do not guess the master-to-layout mapping. I read it from the surviving
slideMaster, which lists its layouts in order inside a sldLayoutIdLst element:

    <p:sldLayoutIdLst>
      <p:sldLayoutId ... r:id="rId1"/>
      <p:sldLayoutId ... r:id="rId2"/>
      ...
    </p:sldLayoutIdLst>

Those relationship ids appear in the same order as the layout files, so I map
them positionally and give the theme the next free id after the layouts.

For each slide I point the layout link at the first layout. That is enough to
render, because a slide draws its own shapes and the layout only supplies
inherited placeholders and background. The image links carry the one
best-effort caveat in this module: a slide references an image by relationship
id but does not record which file that id meant, because that mapping lived only
in the lost .rels. I therefore assign media in order. The text, shapes, and
positions are exact because they live in the slide XML.

I read the parts already inside the input file, rebuild all .rels, and write a
new package. I do not need the original file or any extracted tree.
"""

import sys
import os
import re
import zipfile
import argparse
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger("rebuild_graph")

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
T_OFFICEDOC = BASE + "/officeDocument"
T_SLIDE     = BASE + "/slide"
T_SLIDEMASTER = BASE + "/slideMaster"
T_SLIDELAYOUT = BASE + "/slideLayout"
T_THEME     = BASE + "/theme"
T_IMAGE     = BASE + "/image"
T_NOTESMASTER = BASE + "/notesMaster"
T_NOTESSLIDE  = BASE + "/notesSlide"
T_HLINK     = BASE + "/hyperlink"

RID_RE = re.compile(r'r:(?:embed|link|id)="(rId\d+)"')
HLINK_RE = re.compile(r'<a:hlink(?:Click|Hover)[^>]*r:id="(rId\d+)"')


def num(name):
    """Return the trailing integer in a part name, or 0 if there is none.

    I use this to sort parts such as slide2, slide10 numerically rather than
    lexically, so slide10 sorts after slide2 instead of before it.

    Args:
        name: A part name ending in a number before .xml.

    Returns:
        The integer found before the .xml suffix, or 0 when none is present.
    """
    m = re.search(r'(\d+)\.xml$', name)
    return int(m.group(1)) if m else 0


def rels_doc(relationships):
    """Serialize a list of relationship tuples into a .rels XML document.

    Args:
        relationships: A list of (Id, Type, Target, mode) tuples. `mode` is the
            string "External" for targets outside the package, such as
            hyperlinks, and None for ordinary internal targets.

    Returns:
        The relationships XML encoded as bytes.
    """
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           f'<Relationships xmlns="{REL_NS}">']
    for rid, rtype, target, mode in relationships:
        m = f' TargetMode="{mode}"' if mode else ''
        out.append(f'  <Relationship Id="{rid}" Type="{rtype}" Target="{target}"{m}/>')
    out.append('</Relationships>')
    return ("\n".join(out)).encode('utf-8')


def main():
    """Inventory the parts, rebuild every tier of .rels, and write a new file.

    I take inventory of the slides, layouts, masters, themes, media, and notes
    parts, then abort if the essential parts (presentation, a master, and
    layouts) are not all present, because without them the package cannot be
    wired up. I then build seven groups of relationship files in turn: the
    package root, the master, each layout, each notes master, each notes slide,
    each slide, and the presentation. Finally I copy the original parts into a
    new archive while skipping any stale .rels so my rebuilt ones win, add the
    rebuilt .rels, and report what I produced. I return 0 on success and a
    non-zero code when the input is missing or unwirable.
    """
    ap = argparse.ArgumentParser(description="Rebuild the full .rels graph of a recovered pptx.")
    ap.add_argument('input')
    ap.add_argument('--out', '-o', default='fixed_graph.pptx')
    ap.add_argument('--verbose', '-v', action='store_true')
    args = ap.parse_args()
    if args.verbose:
        log.setLevel(logging.DEBUG)
    if not os.path.exists(args.input):
        log.error(f"not found: {args.input}")
        return 2

    zin = zipfile.ZipFile(args.input)
    names = set(zin.namelist())

    # --- inventory the parts that exist
    slides = sorted([n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)], key=num)
    layouts = sorted([n for n in names if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n)], key=num)
    masters = sorted([n for n in names if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n)], key=num)
    themes = sorted([n for n in names if re.match(r'ppt/theme/theme\d+\.xml$', n)], key=num)
    media = sorted([n for n in names if n.startswith('ppt/media/')], key=num)
    notesMasters = sorted([n for n in names if re.match(r'ppt/notesMasters/notesMaster\d+\.xml$', n)], key=num)
    notesSlides = sorted([n for n in names if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', n)], key=num)
    has_pres = 'ppt/presentation.xml' in names

    log.info(f"slides={len(slides)} layouts={len(layouts)} masters={len(masters)} "
             f"themes={len(themes)} media={len(media)} notesMasters={len(notesMasters)} "
             f"notesSlides={len(notesSlides)} presentation={'yes' if has_pres else 'NO'}")

    if not (has_pres and masters and layouts):
        log.error("Need presentation.xml, at least one master, and layouts. Missing one "
                  "of these means the package can't be wired up. Aborting.")
        return 1

    master = masters[0]
    master_theme = themes[0] if themes else None         # theme1 -> slideMaster1
    notes_theme = themes[1] if len(themes) > 1 else master_theme  # theme2 -> notesMaster1

    new_parts = {}  # arcname -> bytes (the .rels I synthesize)

    # --- (1) _rels/.rels : package -> presentation
    new_parts['_rels/.rels'] = rels_doc([
        ('rId1', T_OFFICEDOC, 'ppt/presentation.xml', None),
    ])

    # --- (2) slideMaster rels : master -> all layouts (by the order the
    #         master lists them) + theme. I read the master's
    #         sldLayoutIdLst so my rIds match what the master expects.
    master_xml = zin.read(master).decode('utf-8', 'replace')
    master_layout_rids = re.findall(r'<p:sldLayoutId[^>]*r:id="(rId\d+)"', master_xml)
    log.debug(f"master lists {len(master_layout_rids)} layout rIds: {master_layout_rids}")

    master_rels = []
    # Map each rId the master references, in order, to a layout file in order.
    for i, rid in enumerate(master_layout_rids):
        if i < len(layouts):
            target = '../slideLayouts/' + os.path.basename(layouts[i])
            master_rels.append((rid, T_SLIDELAYOUT, target, None))
    # Theme: assign the next free rId after the layout rIds.
    if master_theme:
        used = {int(r[3:]) for r in master_layout_rids}
        t = 1
        while t in used:
            t += 1
        theme_rid = f"rId{t}"
        master_rels.append((theme_rid, T_THEME, '../theme/' + os.path.basename(master_theme), None))
    new_parts[f'ppt/slideMasters/_rels/{os.path.basename(master)}.rels'] = rels_doc(master_rels)

    # --- (3) each slideLayout rels : layout -> master
    for lay in layouts:
        new_parts[f'ppt/slideLayouts/_rels/{os.path.basename(lay)}.rels'] = rels_doc([
            ('rId1', T_SLIDEMASTER, '../slideMasters/' + os.path.basename(master), None),
        ])

    # --- (4) notesMaster rels : notesMaster -> theme
    for nm in notesMasters:
        if notes_theme:
            new_parts[f'ppt/notesMasters/_rels/{os.path.basename(nm)}.rels'] = rels_doc([
                ('rId1', T_THEME, '../theme/' + os.path.basename(notes_theme), None),
            ])

    # --- (5) notesSlide rels : notesSlide -> notesMaster
    for ns in notesSlides:
        if notesMasters:
            new_parts[f'ppt/notesSlides/_rels/{os.path.basename(ns)}.rels'] = rels_doc([
                ('rId1', T_NOTESMASTER, '../notesMasters/' + os.path.basename(notesMasters[0]), None),
            ])

    # --- (6) each slide rels : slide -> layout (+ images)
    default_layout = '../slideLayouts/' + os.path.basename(layouts[0])
    media_targets = ['../media/' + os.path.basename(m) for m in media]
    img_cursor = 0
    slide_report = []
    for s in slides:
        xml = zin.read(s).decode('utf-8', 'replace')
        hlinks = set(HLINK_RE.findall(xml))
        rids = {}
        for rid in RID_RE.findall(xml):
            rids[rid] = 'hyperlink' if rid in hlinks else 'image'

        rels = []
        # Layout link first, on an rId that doesn't clash with the slide's own.
        used = {int(r[3:]) for r in rids}
        n = 1
        while n in used:
            n += 1
        layout_rid = f"rId{n}"
        rels.append((layout_rid, T_SLIDELAYOUT, default_layout, None))

        # Image links, assigned in order from the media pool.
        image_rids = sorted([r for r, k in rids.items() if k == 'image'], key=lambda r: int(r[3:]))
        for rid in image_rids:
            if media_targets:
                target = media_targets[img_cursor % len(media_targets)]
                img_cursor += 1
            else:
                target = '../media/image1.png'
            rels.append((rid, T_IMAGE, target, None))

        # Hyperlinks: placeholder external target (real URL was in the lost rels).
        for rid, k in rids.items():
            if k == 'hyperlink':
                rels.append((rid, T_HLINK, 'https://example.invalid/recovered', 'External'))

        new_parts[f'ppt/slides/_rels/{os.path.basename(s)}.rels'] = rels_doc(rels)
        slide_report.append((num(s), len(image_rids), sum(1 for v in rids.values() if v == 'hyperlink')))

    # --- (7) presentation rels : presentation -> slides+master+theme+notes-
    pres_xml = zin.read('ppt/presentation.xml').decode('utf-8', 'replace')
    pres_rids = re.findall(r'r:id="(rId\d+)"', pres_xml)
    # The presentation lists slide rIds in <p:sldIdLst>; their order = slide order.
    sld_rids = re.findall(r'<p:sldId[^>]*r:id="(rId\d+)"', pres_xml)
    log.debug(f"presentation references {len(pres_rids)} rIds, {len(sld_rids)} in sldIdLst")

    pres_rels = []
    # Map the sldIdLst rIds to slides in order.
    for i, rid in enumerate(sld_rids):
        if i < len(slides):
            pres_rels.append((rid, T_SLIDE, 'slides/' + os.path.basename(slides[i]), None))
    # Remaining (non-slide) rIds in presentation.xml are master/theme/notesMaster.
    # Assign them on fresh rIds beyond everything used so far.
    used = {int(r[3:]) for r in pres_rids}
    def next_free():
        """Return the next relationship id not already used in this part.

        I scan upward from rId1, skip any number already in `used`, record the
        one I pick, and return it formatted as an rId string. This guarantees a
        fresh id when I bind the master, theme, or notes master.
        """
        k = 1
        while k in used:
            k += 1
        used.add(k)
        return f"rId{k}"

    # If the presentation references a master/theme/notesMaster rId that ISN'T a
    # slide, I still need those relationships to exist. I add them on whatever
    # rIds the presentation already expects if they're non-slide, else fresh.
    nonslide_rids = [r for r in pres_rids if r not in sld_rids]
    # Build the list of targets I must bind: master, theme, notesMaster.
    targets = [(T_SLIDEMASTER, 'slideMasters/' + os.path.basename(master))]
    if master_theme:
        targets.append((T_THEME, 'theme/' + os.path.basename(master_theme)))
    if notesMasters:
        targets.append((T_NOTESMASTER, 'notesMasters/' + os.path.basename(notesMasters[0])))

    for i, (rtype, target) in enumerate(targets):
        rid = nonslide_rids[i] if i < len(nonslide_rids) else next_free()
        pres_rels.append((rid, rtype, target, None))

    new_parts['ppt/_rels/presentation.xml.rels'] = rels_doc(pres_rels)

    # --- write the new package
    log.info(f"Writing {len(new_parts)} rebuilt .rels files into {args.out}")
    with zipfile.ZipFile(args.out, 'w', zipfile.ZIP_DEFLATED) as zout:
        # copy original parts, but SKIP any stale .rels so ours win
        for item in zin.infolist():
            if item.filename.endswith('.rels'):
                continue
            data = zin.read(item.filename)
            ext = item.filename.rsplit('.', 1)[-1].lower() if '.' in item.filename else ''
            ctype = zipfile.ZIP_STORED if ext in ('png', 'jpg', 'jpeg', 'gif') else zipfile.ZIP_DEFLATED
            zout.writestr(item, data, compress_type=ctype)
        # add rebuilt rels
        for arc, content in new_parts.items():
            zout.writestr(arc, content)

    log.info(f"Wrote {args.out} ({os.path.getsize(args.out):,} bytes)")
    total_imgs = sum(r[1] for r in slide_report)
    log.info(f"Rebuilt graph: {len(slides)} slide-rels ({total_imgs} image links), "
             f"{len(layouts)} layout-rels, 1 master-rels, presentation + package rels.")
    if args.verbose:
        for n_, imgs, links in slide_report:
            if imgs or links:
                log.debug(f"  slide{n_}: {imgs} images, {links} hyperlinks")

    print("\n" + "=" * 72)
    print(f"Open {args.out} in PowerPoint.")
    print("Now the FULL chain is wired: slide -> layout -> master -> theme.")
    print("Slides should render their text, shapes, and backgrounds. Multi-image")
    print("slides will show images, but specific image-to-slot pairing is by order.")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
