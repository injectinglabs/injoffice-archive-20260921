#!/usr/bin/env python3
"""Independent validation of S12's diagram-slice connector fields
(HeadArrow/TailArrow/FlipH -> a:headEnd/a:tailEnd/@flipH) against
DiagramExampleDeck's output (cmd/diagram_sample).

python-pptx (1.0.2) has NO friendly Python property for connector
arrowheads or xfrm flip — grep the oxml module and the only place
a:headEnd/a:tailEnd/flipH appear is in internal child-ordering "successors"
lists (CT_LineProperties, CT_Transform2D), never as a readable attribute.
So this checks the real thing python-pptx DOES give independent access to:
each shape's raw XML as parsed by lxml (shape._element.xml) — a completely
separate XML parser from pptxpatch's own encoding/xml writer, still proof
the elements/attributes are well-formed and exactly where intended, just
not through a high-level property.

Usage: python3 scripts/validate_diagram_arrows.py <path-to-generated-diagram.pptx>
"""
import re
import sys

from pptx import Presentation


def main():
    if len(sys.argv) != 2:
        print("usage: validate_diagram_arrows.py <path.pptx>", file=sys.stderr)
        sys.exit(2)
    path = sys.argv[1]
    prs = Presentation(path)
    errors = []

    def find(name, slide_idx):
        for sh in prs.slides[slide_idx].shapes:
            if sh.name == name:
                return sh
        errors.append(f"slide {slide_idx + 1}: no shape named {name!r}")
        return None

    def xml_of(shape):
        return shape._element.xml  # noqa: SLF001 — deliberate raw-XML escape hatch, see module doc

    # --- slide 1: process-flow arrows (TailArrow on every connector) ---
    for name in ("a1", "a2", "a3"):
        sh = find(name, 0)
        if sh is None:
            continue
        xml = xml_of(sh)
        if "tailEnd" not in xml:
            errors.append(f"{name}: expected <a:tailEnd> (TailArrow), not found in:\n{xml}")
        if 'type="triangle"' not in xml:
            errors.append(f"{name}: expected a triangle arrowhead type")
        if "headEnd" in xml:
            errors.append(f"{name}: unexpected <a:headEnd> — this connector should be tail-only")
        # tailEnd must come AFTER solidFill in document order (CT_LineProperties'
        # schema sequence) — a real ordering bug here is exactly the kind of
        # thing that opens fine in a lenient reader but PowerPoint itself
        # rejects as a corrupt file.
        ln_match = re.search(r"<a:ln\b.*?</a:ln>", xml, re.S)
        if ln_match:
            ln = ln_match.group(0)
            fill_idx, tail_idx = ln.find("solidFill"), ln.find("tailEnd")
            if fill_idx != -1 and tail_idx != -1 and tail_idx < fill_idx:
                errors.append(f"{name}: <a:tailEnd> appears before <a:solidFill> — invalid child order")

    # --- slide 2: org-chart connectors (plain lines, no arrows) ---
    for name in ("c-left", "c-mid", "c-right"):
        sh = find(name, 1)
        if sh is None:
            continue
        xml = xml_of(sh)
        if "headEnd" in xml or "tailEnd" in xml:
            errors.append(f"{name}: org-chart connector should carry NO arrowheads, found one in:\n{xml}")

    # flipH: only c-left should carry it (child left of parent center); c-mid
    # (dx=0, degenerate) and c-right (already the main diagonal) should not.
    left = find("c-left", 1)
    mid = find("c-mid", 1)
    right = find("c-right", 1)
    if left is not None and 'flipH="1"' not in xml_of(left):
        errors.append("c-left: expected flipH=\"1\" on its <a:xfrm> (anti-diagonal case), not found")
    if mid is not None and 'flipH="1"' in xml_of(mid):
        errors.append("c-mid: unexpected flipH on a degenerate (dx=0) connector")
    if right is not None and 'flipH="1"' in xml_of(right):
        errors.append("c-right: unexpected flipH — this connector is already the main diagonal")

    if errors:
        print("VALIDATION FAILED:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)

    print(f"OK: {path} — diagram connector arrowheads (head/tail) and flipH all independently verified via python-pptx's raw XML (lxml-backed, not our own writer/reader).")


if __name__ == "__main__":
    main()
