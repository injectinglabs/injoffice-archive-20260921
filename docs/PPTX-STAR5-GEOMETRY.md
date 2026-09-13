# Default five-point star geometry

`presetPath('star5', widthEmu, heightEmu)` and authored PPTX rendering now use
DrawingML's default five-point star guides. The authored compiler already
accepts `kind: 'star5'`; the playground's **Native shape preset** selector on
its PPTX render page displays the corrected path.

The default guides use `adj=19098`, `hf=105146`, and `vf=110557`. Outer vertices
share the pentagon's frame scaling; inner vertices use the preset adjustment.
Coordinates retain guide precision until rounding once to integer EMU. Explicit
`textBody` layout uses the preset rectangle (`sx1`, `sy1`, `sx4`, `sy3`) before
body insets. Frame rotation then applies to both shape and text bounds.

This corrects the former generic inscribed star, so existing authored star
outlines, text placement, and render hashes change. Legacy paragraph placement
without explicit text-body metadata keeps its previous frame-bound behavior.
Tiny geometry can collapse after integer rounding; insets that collapse a text
region are refused by the existing compiler bounds check.

This is a default authored-geometry correction. It adds no source extraction,
custom adjustment, custom-path, font, or Microsoft Office visual-equivalence
qualification. Source stars remain subject to existing extraction refusals.

Guide reference: [Apache POI's published DrawingML preset definitions](https://github.com/apache/poi/blob/trunk/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml), `star5`.
