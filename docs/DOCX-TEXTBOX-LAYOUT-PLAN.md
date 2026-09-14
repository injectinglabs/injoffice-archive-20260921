# DOCX textbox layout delivery plan

This tracks the remaining textbox work requested after PR #174. Footnote
continuation, multiple page-placed rectangles, and ordinary relative offsets
and alignment are already merged. Each item below requires source-bound
inspection, layout/composition, browser presentation, positive and refusal
regressions, real-DOCX evidence, green CI, and merge before it is delivered.

- [x] Stacking (merged in PR #179): honor front/behind-body placement and unsigned relative height;
  preserve source order for ties; prove overlap pixels and reject forged layers.
- [x] Parity-relative margins (merged in PR #182): resolve authored inside/outside and physical margin
  bases and alignments on odd/even pages, with source section geometry.
- [x] Anchors after text (merged in PR #186): locate the source drawing's actual
  shaped line and character position rather than always using the paragraph's first line.
- [x] Inline textboxes (merged in PR #188): reserve inline width and line height,
  reflow surrounding text, paginate the object with its owning line, and paint its exact shape;
  malformed or unsupported source shapes remain preserve-only.
- [x] Body wrapping (merged in PR #192): reserve source-qualified page-placed rectangles
  in paragraph line layout, displace fully blocked lines across page transitions, and retain
  explicit refusals for interior/overlapping rectangles, numbering, and unsupported wrap modes.

Completion means these supported behaviors are implemented and validated, rather than
marking their refusals as completed functionality. Source documents must remain
unchanged and original drawing diagnostics retained in the read-only result.
The preview remains approximate: top-and-bottom and other unsupported wrap geometries
continue to refuse with source diagnostics, and this plan does not claim unrestricted
Word layout equivalence for unrelated unsupported document features.
