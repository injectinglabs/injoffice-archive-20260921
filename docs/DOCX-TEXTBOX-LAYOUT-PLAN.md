# DOCX textbox layout coverage

The native preview supports the following source-qualified textbox behaviors.
Regression coverage includes extraction, layout, browser presentation, and
explicit refusal of unsupported source shapes.

- **Stacking:** honor front/behind-body placement and unsigned relative height;
  preserve source order for ties; prove overlap pixels and reject forged layers.
- **Parity-relative margins:** resolve authored inside/outside and physical margin
  bases and alignments on odd/even pages, with source section geometry.
- **Anchors after text:** locate the source drawing's actual
  shaped line and character position rather than always using the paragraph's first line.
- **Inline textboxes:** reserve inline width and line height,
  reflow surrounding text, paginate the object with its owning line, and paint its exact shape;
  malformed or unsupported source shapes remain preserve-only.
- **Body wrapping:** reserve source-qualified page-placed rectangles
  in paragraph line layout, displace fully blocked lines across page transitions, and retain
  explicit refusals for interior/overlapping rectangles, numbering, and unsupported wrap modes.

Source documents remain unchanged and original drawing diagnostics are retained
in the read-only result. Supported behavior does not make refused geometry editable.
The preview remains approximate: top-and-bottom and other unsupported wrap
geometries continue to refuse with source diagnostics. This coverage does not
claim unrestricted Word layout equivalence.
