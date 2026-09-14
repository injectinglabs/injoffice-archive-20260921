# Graphic-frame qualification preparation

This is dedicated preparation, not new source admission. Newly complex table
and chart transforms retain their existing refusal guards until the complete
source/compiler/worker/browser path and reference policy are qualified.

The renderer must qualify all painted control hulls under the final world
mapping, including chart labels outside the plot, tick/stroke envelopes, table
glyphs and refusal placeholders. Checking only the frame rectangle is insufficient.
`sourceGraphicFramePolicy` decodes existing paint transforms without replacing
rational coefficients by rounded binary64 values. It carries conversion errors
through composition and charges the shared request budget. A frame permits at
most 8,192 explicitly supplied spans; sparse lists refuse.

`sourceTableTextBounds` follows the actual text paint order: body orientation,
writing mode, paragraph placement, then supplied glyph or refused-run placeholder.
A wholly refused text body paints its untransformed body placeholder instead.
The helper retains cell x clipping edges and encloses the complete local y hull,
then qualifies both text and clip corners under the actual cell world mapping.
Backgrounds and borders must remain outside this text-only clip. Callers must
supply a complete actual world affine and measured glyph extents; no font metrics
or source authority are invented by these helpers.

## Reference matrix and remaining policy decision

PresentationML graphicFrame/xfrm uses DrawingML CT_Transform2D. An atomic chart
surface can rotate/scale its plot and labels coherently, but an odd reflection
would mirror glyphs unless a separately qualified chart text policy intervenes.
Shape/table counter-reflection must not be silently applied to chart labels.

An external matrix retains source hashes for table/chart baseline, R30, H, V,
and a nested R30/H group with horizontal scale 1.5 and child R90. The normalized
Office-openable chart container is preserved; only slide1.xml changes. The table
uses asymmetric labels and distinct colors. The coordinator's sole Office owner
exports isolated source copies and compares saved XML before pixel/coordinate
interpretation. These references will determine the disclosed text policy before
shared source guards change. No Office parity claim follows from helper tests.

Required connected evidence includes rotated/reflected/anisotropic actual font
outlines, an outside-cell-x negative control, retained outside-cell-y overflow,
active slide clipping, complete world/error-budget failures, immutable source,
and preview-only mutation restrictions. Shared hooks follow the bubble-chart
integration handoff; this helper preparation is not a standalone completion PR.

### Direct-frame implementation difference

[MS-OE376 §2.1.1365(b)](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/0967224b-efc3-42dd-907a-3f8c1a240172)
explicitly documents that Office ignores graphic-frame `flipH`, `flipV` and
`rot`. This differs from the standard transform description. The external
PowerPoint16.112.4 reference independently corroborates it: direct baseline,
R30, H and V produce byte-identical PNGs for both tested families, while their
source transform attributes survive the Office save. All chart graph bytes are
unchanged. This supports a named Microsoft implementation profile for direct
frames; it does not qualify arbitrary ancestor-group layout.

Grouped chart references retain upright glyph metrics and relayout the chart
surface at physical dimensions. A child quarter-turn changes group scale-axis
selection even though its final chart surface stays axis-aligned. Grouped table
references retain intrinsic grid/row physical sizes and rewrite frame/group
anchors on Office save. Consequently, blindly applying a generic affine to all
existing table glyphs would also preserve an earlier parsed-group discrepancy.
These distinct anchor, relayout and source-normalization behaviors remain a
required integration task. Offset-child references discriminate ancestor center
mapping independently from the centered fixtures.

## Physical anchor preparation

The internal `office-graphic-frame-anchor-v1` helper now keeps physical width and height as bounded rationals and emits a translation-only origin. It maps the source center conventionally through nearest-first ancestors. Positive ancestor scale products size the physical anchor; the leaf's raw rotation selects swapped scale axes in the half-open intervals [45°,135°) and [225°,315°). Direct orientation does not rotate or reflect the painted surface. This is an explicitly named implementation policy, not a replacement for standard DrawingML semantics.

The retained 32-case Office 16.112.4 reference matrix includes direct rotation/reflections, offset children, group rotation/reflection/nonuniform scaling, and 44°/45°/134°/135° discriminators. Sixteen table vector bounds agree with the independently calculated anchor plus intrinsic tracks within 23.4375 EMU, consistent with the PDF export's vector quantization. That observation is separate from, and does not loosen, the renderer's 0.125 EMU conversion allowance. Office normalizes grouped table transforms on save; original and normalized source transforms are retained separately.

This remains preparation: callers must qualify the final complete paint hulls and preserve source authority. Tables use intrinsic row/column dimensions at the resulting origin. Charts require physical layout, not scaling already laid-out glyphs. Fractional physical chart dimensions still need an exact layout adapter or an explicitly reviewed host quantization policy before admission. Arbitrarily nested noncommuting groups require additional discriminating reference coverage before this profile is connected broadly.
