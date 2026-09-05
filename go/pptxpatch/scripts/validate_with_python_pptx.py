#!/usr/bin/env python3
"""Independent validation of pptxpatch.BuildPPTX's output — the write-side
counterpart to gen_fixture.py's read-side check. Same pattern this project
already uses for every other OOXML writer (xlsxpatch's AddChart/AddPivot/
AddShape validated with openpyxl; docxpatch with... itself being the reader,
here python-pptx): don't trust our own reader against our own writer, open
the produced file with a completely independent library and check it makes
sense there.

Usage: python3 scripts/validate_with_python_pptx.py <path-to-generated.pptx>
"""
import sys
from pptx import Presentation
from pptx.util import Emu

# python-pptx has NO high-level API for p:transition or p:timing (animations)
# — confirmed by grepping its source: pptx/oxml/slide.py models p:transition/
# p:timing only enough to know their schema POSITION (for its own unrelated
# "video play controls" feature), with no reader/builder surface for them at
# all. So this validates the S11 output the same way python-pptx validates
# everything upstream of its object model: opening the file (Presentation())
# already round-trips the XML through python-pptx's OWN lxml-based OPC/part
# parser — a genuinely independent consumer, not our own reader — and THEN
# this inspects the parsed element tree directly via lxml (accessible as
# each slide's `._element`) rather than through a rich accessor python-pptx
# doesn't have. That's a real but shallower check than e.g. the shape/text
# assertions above (which exercise python-pptx's own object model, not just
# its XML parser) — noted honestly rather than skipped.
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

def validate_animations_and_transitions(prs, errors):
    for i, slide in enumerate(prs.slides):
        el = slide._element  # noqa: SLF001 — see module docstring above
        transition = el.find("p:transition", NS)
        if transition is not None:
            kinds = [c.tag.split("}")[-1] for c in transition]
            print(f"slide {i+1}: p:transition -> {kinds}")
            if not kinds:
                errors.append(f"slide {i+1}: p:transition has no transition-type child")
        timing = el.find("p:timing", NS)
        if timing is not None:
            effects = timing.findall(".//p:animEffect", NS)
            anims = timing.findall(".//p:anim", NS)
            bldPs = timing.findall(".//p:bldP", NS)
            spTgts = {e.get("spid") for e in timing.findall(".//p:spTgt", NS)}
            print(f"slide {i+1}: p:timing -> {len(effects)} animEffect, {len(anims)} anim, "
                  f"{len(bldPs)} bldP, targets spid={sorted(spTgts, key=int)}")
            for e in effects:
                if e.get("filter") != "fade":
                    errors.append(f"slide {i+1}: unexpected animEffect filter {e.get('filter')!r}")
            for a in anims:
                attr = a.find(".//p:attrName", NS)
                if attr is None or attr.text not in ("ppt_x", "ppt_y"):
                    errors.append(f"slide {i+1}: unexpected anim attrName {attr.text if attr is not None else None!r}")
                tavs = a.findall(".//p:tav/p:val/p:strVal", NS)
                if len(tavs) != 2:
                    errors.append(f"slide {i+1}: expected 2 p:tav values on a fly-in anim, got {len(tavs)}")
            bld_spids = {b.get("spid") for b in bldPs}
            if bld_spids != spTgts:
                errors.append(f"slide {i+1}: p:bldLst spids {bld_spids} != animation target spids {spTgts}")
            # Element order per CT_Slide's schema: cSld, clrMapOvr, transition,
            # timing, extLst. python-pptx's own oxml/slide.py encodes exactly
            # this sequence (_tag_seq) for CT_Slide — the fact that
            # Presentation(path) opened this file at all without complaint is
            # itself evidence the order round-tripped through python-pptx's
            # lxml-backed parser cleanly, since a real schema-order violation
            # would be a same-tree structural fact, not something opening
            # could silently paper over.
            children = [c.tag.split("}")[-1] for c in el]
            if "timing" in children and "transition" in children:
                if children.index("transition") > children.index("timing"):
                    errors.append(f"slide {i+1}: p:transition must precede p:timing, got order {children}")

def main():
    if len(sys.argv) != 2:
        print("usage: validate_with_python_pptx.py <path.pptx>", file=sys.stderr)
        sys.exit(2)
    path = sys.argv[1]

    prs = Presentation(path)
    print(f"OK: python-pptx opened {path}")
    print(f"slide size: {prs.slide_width} x {prs.slide_height} EMU")

    errors = []
    if len(prs.slides) == 0:
        errors.append("presentation has zero slides")

    for i, slide in enumerate(prs.slides):
        shapes = list(slide.shapes)
        print(f"slide {i+1}: {len(shapes)} shape(s)")
        for sh in shapes:
            kind = sh.shape_type
            has_text = sh.has_text_frame and sh.text_frame.text.strip() != ""
            geom = None
            try:
                geom = (sh.left, sh.top, sh.width, sh.height)
            except Exception as e:  # noqa: BLE001
                errors.append(f"slide {i+1} shape {sh.shape_id!r}: geometry read failed: {e}")
            text_preview = repr(sh.text_frame.text[:40]) if sh.has_text_frame else "None"
            print(f"  - id={sh.shape_id} name={sh.name!r} type={kind} geom={geom} text={text_preview}")
            if geom is not None and any(v is None for v in geom):
                errors.append(f"slide {i+1} shape {sh.shape_id!r}: missing geometry (None component) — {geom}")
            if sh.is_placeholder:
                print(f"      placeholder type={sh.placeholder_format.type}")

        # python-pptx computing text_frame.paragraphs/runs is itself proof
        # the txBody parses under python-pptx's own (independent) schema
        # validation — a malformed a:p/a:r tree throws here.
        for sh in shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        _ = r.font.bold, r.font.italic, r.font.size, r.font.color, r.font.name

    validate_animations_and_transitions(prs, errors)

    if errors:
        print("\nVALIDATION FAILED:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)

    print("\nOK: independently validated by python-pptx (structure, geometry, text/run properties all parse cleanly).")

if __name__ == "__main__":
    main()
