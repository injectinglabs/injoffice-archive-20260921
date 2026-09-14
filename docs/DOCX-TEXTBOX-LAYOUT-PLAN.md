# DOCX textbox layout delivery plan

This tracks the remaining textbox work requested after PR #174. Footnote
continuation, multiple page-placed rectangles, and ordinary relative offsets
and alignment are already merged. Each item below requires source-bound
inspection, layout/composition, browser presentation, positive and refusal
regressions, real-DOCX evidence, green CI, and merge before it is delivered.

- [x] Stacking (merged in PR #179): honor front/behind-body placement and unsigned relative height;
  preserve source order for ties; prove overlap pixels and reject forged layers.
- [ ] Parity-relative margins: resolve authored inside/outside and physical margin
  bases and alignments on odd/even pages, with source section geometry.
- [ ] Anchors after text: locate the source drawing's actual shaped line and
  character position rather than always using the paragraph's first line.
- [ ] Inline textboxes: reserve inline width and line height, reflow surrounding
  text and paginate the object with its owning line, then paint its exact shape.
- [ ] Body wrapping: reserve source-qualified floating rectangles in paragraph
  line layout, including square and top/bottom wrapping and page transitions.

Completion means these behaviors are implemented and validated, rather than
marking their refusals as completed functionality. Source documents must remain
unchanged and original drawing diagnostics retained in the read-only result.
The preview remains approximate: this plan does not claim unrestricted Word
layout equivalence for unrelated unsupported document features.
