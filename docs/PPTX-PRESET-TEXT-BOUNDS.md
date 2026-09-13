# Ellipse, triangle, and diamond text regions

Authored and already-qualified parsed shapes with explicit `textBody` metadata
now lay out text in the DrawingML preset rectangle before applying body insets.
Previously these three presets incorrectly used the full rectangular frame.

| Preset | Left | Top | Right | Bottom |
| --- | --- | --- | --- | --- |
| ellipse | w/2 − w/(2√2) | h/2 − h/(2√2) | w/2 + w/(2√2) | h/2 + h/(2√2) |
| triangle, default adj=50000 | w/4 | h/2 | 3w/4 | h |
| diamond | w/4 | h/4 | 3w/4 | 3h/4 |

Edges round once to integer EMU before body insets are applied. Text wrapping,
anchors, and quarter-turn transforms use the resulting inner region. Insets
that fit the full frame but collapse the preset region are refused. Tiny regions
can collapse after integer rounding and are subject to the same bounds check.

The existing `compileWireDeckToNativeV1` and `compileNativePptxSlide` APIs expose
this correction without a new option. Shape outlines and source qualification
are unchanged. Legacy paragraph placement with no explicit `textBody` retains
its full-frame behavior. This is a geometry correction; it does not qualify
fonts, arbitrary adjustments, or Microsoft Office visual equivalence.

Guide reference: [Apache POI's published DrawingML preset definitions](https://github.com/apache/poi/blob/trunk/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml), `ellipse`, `triangle`, `diamond`.
