# DrawingML source text orientation

Source extraction, the public compiler, recording paint and the font worker now
preview separate `bodyPr@rot` and `upright=true` for text boxes and shapes.
The native contract preserves raw signed `rotationAngle60000` and explicit
`upright` presence. Nonzero rotation/upright parsed targets remain preserve-only,
and actual mutation resolution rejects them. Existing absent/default behavior
is unchanged. This is a bounded deterministic preview, not Office layout parity.

## Source rules and preview policy

ECMA-376 Part 1, `bodyPr` (21.1.2.1.1), distinguishes text rotation from shape
rotation. `rot` is a signed 32-bit `ST_Angle` in 1/60000-degree units; its default
is zero. The parser retains the exact signed source value and attribute presence.
Only trigonometric evaluation reduces complete turns. Canonical numeric syntax,
strict boolean spelling, duplicate attributes and foreign names are checked.

The [Microsoft rotation example](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rotation?view=openxml-3.0.1)
combines a +90-degree shape with −90-degree body rotation to leave its text
unrotated. A permanent nonsquare-frame test verifies the exact identity matrix.

The body-rotation helper follows the existing pinned Apache POI text preview
policy: accumulated odd reflection receives a local horizontal counter-reflection,
then body rotation acts about the group-scaled anchor center, before writing-mode
layout. This ordering is supported by
[DrawTextShape](https://github.com/apache/poi/blob/338882ac8898df5c13a7d15f533204c5dd8607d6/poi/src/main/java/org/apache/poi/sl/draw/DrawTextShape.java)
and its group-aware `DrawShape.getAnchor` call. This is a disclosed deterministic
preview policy, not a PowerPoint fidelity claim.

For positive own-axis group scale `S`, body rotation `B`, and existing text
counter-reflection `C`, the additional local transform is `C S⁻¹ B S`. The node
already supplies the source hierarchy transform `R F S`. Rotating directly
below a nonuniform scale would instead distort glyph axes. Centered scales use
reduced exact rationals, with the existing 512-bit limit. All operations share
the compile request budget; noncardinal angles retain the affine engine's
explicit uncertainty allowance. The caller must qualify complete glyph/control
and clip hulls after composition, including text outside the original frame.
No formal libm enclosure or Office parity is claimed.

## Upright area proposal (dedicated helper only)

The [Microsoft upright example](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.upright?view=openxml-3.0.1)
establishes that upright text ignores both shape and body rotation. It does not
fully specify nonsquare text-area sizing under grouped nonuniform scales.
Simply cancelling the shape angle after the existing layout is therefore not
a complete implementation.

Pinned [LibreOffice shape import](https://github.com/LibreOffice/core/blob/c0df28f957b43f0ecbffdc994ee0702dbe5bc33f/oox/source/drawingml/shape.cxx#L1643)
uses distinct text-area and text-pre-rotation concepts. Its near-vertical
quadrants, [45°,135°) and [225°,315°), rotate the area by a quarter turn and
counterrotate text. This supports a named preview policy with swapped physical
text-area dimensions in those quadrants, rather than a claim that every detail
is normatively prescribed.

The dedicated `sourceUprightTextArea` helper separates the original immutable
content rectangle from a derived exact rational layout area. Its center is
preserved, including half-EMU offsets. Under scales `sx,sy`, the swapped layout
dimensions are `(sy/sx)·height` and `(sx/sy)·width`: glyph scales remain `sx,sy`.
Swapping the glyph scales to obtain integer dimensions would alter the supplied
font metrics and is deliberately avoided. Body rotation is
ignored when upright is true. Writing-mode direction, alignment, clipping and
glyph orientation remain separate decisions. Exact quadrant boundaries, all
flips, nonuniform groups, asymmetric insets, real glyph ink and overflow controls
are covered by independently reviewed exact helpers and connected supplied-font
tests. Rational area bounds are not rounded into integer layout rectangles:
the layout adapter compares wrap thresholds and computes anchor offsets exactly,
then qualifies final paint conversion. Existing source dimensions and insets
remain immutable.

### Connected layout representation

`sourceTextLayoutRational` now provides bounded exact margin subtraction, wrap
comparisons and alignment/anchor offsets using the affine engine's existing
rational arithmetic. The paragraph transport retains integer run positions and stores fractional
line-alignment and area-origin residuals in a paragraph-local affine translation.
Glyph metrics are not rescaled and the area is not rounded before wrapping. The unchanged integer metrics and advances
remain authoritative for shaping; a fractional origin belongs to placement.

Current admitted shape/text bodies permit overflow in both axes, so their exact
rational area controls wrapping and placement without introducing a new body
clip. The existing slide clip remains active. Outward integer area envelopes
are used only for bounds qualification. Every paragraph's actual glyph and
refusal-placeholder hull is checked under its placement, text orientation and
final world matrix before painting.

The prepared exact clip/inverse helper is reserved for future independently
qualified fractional clip policies; it is not substituted for source overflow.
New orientation inside authored conventional nontranslation groups explicitly
refuses unless the caller uses a source-mode hierarchy. Existing authored group
behavior without new orientation fields remains unchanged.

Actual functional fixtures exercise Go and WASM extraction, public compilation,
supplied DejaVu Sans glyph outlines and the existing browser vector component:
+90 shape/−90 body, a fractional upright area under 1.5× horizontal group scale,
and reflected/noncardinal upright text. A separate above-slide upright fixture
retains glyph paths but paints zero glyph pixels under the active slide clip.
Independent tests also cover raw signed limits, explicit zero/false presence,
immutable source bytes and actual mutation refusal.

## Required graphic-frame integration

PresentationML `graphicFrame/xfrm` uses DrawingML `CT_Transform2D`, so its own
rotation/flips and transformed group context both require qualification. Current
guards deliberately refuse newly complex table/chart frames atomically.

Tables require cell-local text orientation and full glyph-hull qualification
under the final world matrix. Their current horizontal clip implementation is
explicitly limited to a top-level unrotated table. Replacing that guard requires
a qualified local clip that preserves vertical overflow; dropping the guard or
clipping every cell to its full rectangle would change behavior.

The dedicated `sourceTableTextClip` helper constructs a finite cell-local strip
with unchanged cell x boundaries and a vertical span enclosing every measured
glyph/control and refusal-placeholder hull. It retains the existing slide clip.
The local affine qualifier must pass before outward rounding includes its full
1/8-EMU allowance; the caller must additionally qualify the returned strip under
the final world transform. Actual rotated/flipped/anisotropic browser controls
must prove outside-x ink disappears while outside-y ink remains visible before
source admission changes.

Charts require equivalent source-label orientation and full label/plot/control
hulls, including axis text and its existing supplied-font rules. These source,
compiler and paint changes will follow the chart integration ownership handoff.
Authored legacy transform semantics and parsed-source mutation restrictions must
remain unchanged. A connected source/WASM/worker/browser proof is required;
dedicated helpers alone are not a shippable completion milestone.
