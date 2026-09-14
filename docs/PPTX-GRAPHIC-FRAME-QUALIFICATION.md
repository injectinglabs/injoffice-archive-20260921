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

PresentationML graphicFrame/xfrm uses DrawingML CT_Transform2D. The qualified
source-anchor policy below keeps table/chart surfaces axis-aligned and preserves
supplied glyph metrics; it does not apply shape counter-reflection to chart text.

The external matrix retains original and Office-saved table/chart sources, frame
fill discriminators, hashes, and vector/raster outputs. The coordinator's sole
Office owner exported isolated copies and compared source transforms separately
from their rendered positions. The resulting source-anchor policy and known
Office importer difference are detailed below. No Office parity claim follows
from helper tests alone.

Required connected evidence includes rotated/reflected/anisotropic actual font
outlines, an outside-cell-x negative control, retained outside-cell-y overflow,
active slide clipping, complete world/error-budget failures, immutable source,
and preview-only mutation restrictions. Shared hooks follow the serialized chart
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

The internal `source-graphic-frame-anchor-v1` helper now keeps physical width and height as bounded rationals and emits a translation-only origin. It accepts nearest-first ancestors and normalizes their physical anchors from the outermost group inward (see the nested-reference correction below). Positive inherited scale products size the physical anchor; the leaf's raw rotation selects swapped scale axes in the half-open intervals [45°,135°) and [225°,315°). Direct orientation does not rotate or reflect the painted surface. This is an explicitly named implementation policy, not a replacement for standard DrawingML semantics.

The retained 32-case Office 16.112.4 reference matrix includes direct rotation/reflections, offset children, group rotation/reflection/nonuniform scaling, and 44°/45°/134°/135° discriminators. Sixteen table vector bounds agree with the independently calculated anchor plus intrinsic tracks within 23.4375 EMU, consistent with the PDF export's vector quantization. That observation is separate from, and does not loosen, the renderer's 0.125 EMU conversion allowance. Office normalizes grouped table transforms on save; original and normalized source transforms are retained separately.

This remains preparation: callers must qualify the final complete paint hulls and preserve source authority. Tables use intrinsic row/column dimensions at the resulting origin. Charts require physical layout, not scaling already laid-out glyphs. Fractional physical chart dimensions use the separately named projection preparation described below; its hull and source-authority obligations must be connected before admission. The nested source policy and its explicit importer limitation are detailed below; connected admission remains a separate reviewed step.

The dedicated `nearest-emu-physical-layout-v1` projection now provides a bounded integration option for integer chart layout APIs. It rounds each positive exact physical dimension once (nearest EMU, positive ties upward), returns its exact rational absolute error at most 0.5 EMU, and refuses a result of zero or a result above the caller's coordinate budget. The exact source dimensions and anchor remain separate. Callers must lay out directly in that projected frame without scaling glyphs and qualify complete output hulls with the returned outward size allowances in addition to affine conversion allowances. Its fit, overlap, and wrapping decisions belong to the projected frame; no continuity or equivalence to unquantized layout is asserted. Source-only metadata, a diagnostic, and actual mutation refusal remain required at integration.

## Nested source anchors and Office import differences

The initial conventional ancestor composition failed six nested discriminators and was replaced. The corrected helper normalizes each group's physical anchor outer-to-inner, carries rotation/reflection orientation separately, and swaps inherited positive scale axes at each group's raw quadrant before descending. All three independently filled nested chart frames agree at their exported vector edges. All 19 Office-retained table sources agree with their vector backgrounds within 23.6875 EMU export quantization. Source nesting is bounded separately from temporary arithmetic compositions, which still charge all operations and propagate coefficient errors.

**This source-anchored policy does not emulate PowerPoint's table importer rewrite.** Two original synthetic noncardinal nested tables import at different positions. Relative to the source model, the retained turn/reflection case differs by about +11,673 EMU horizontally and −85,791 EMU vertically; the double-turn case differs by about −29,644 EMU horizontally and −68,429 EMU vertically. Exporting the first case before any save and after normalized PPTX save produced identical raster hashes: this is import behavior, not export ordering. PowerPoint's saved table/group transforms differ from the original, and the corrected source model matches those saved transforms' rendered positions.

The preview must preserve the exact original frame and intrinsic table tracks, not infer or apply that importer rewrite. Original-vs-normalized comparisons remain retained qualification evidence; they are not relabeled as successful Office parity tests. Permanent tests instead assert independently derived source-model invariants and actual Office-retained source cases. Product admission still requires the explicit source-only profile, read-only mutation guard, complete ink/control/clip hull checks, and connected browser evidence.

Microsoft's [grouping support documentation](https://support.microsoft.com/en-gb/office/graphics-visuals/group-or-ungroup-shapes-pictures-or-other-objects) says the Group command is unavailable when the selection contains a table, including in the macOS instructions. This is UI behavior, not a prohibition in the source XML grammar and not a reason to drop grouped-table support. It reinforces keeping the imported table normalization evidence distinct from chart grouping behavior.
