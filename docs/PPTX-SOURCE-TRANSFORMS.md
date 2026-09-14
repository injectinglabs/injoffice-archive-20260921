# DrawingML source transform qualification

Native source extraction, the public compiler, recording paint, and the native
slide paint worker support bounded arbitrary shape/text-box rotation, reflection,
and rational group coordinates. These are preview operations. The source bytes,
object identity, and mutation authority remain separate from rendered geometry.

The native transform adds `rotationAngle` (0 through 21,599,999, in 60,000 units per
degree), `flipH`, and `flipV`. These fields cannot coexist with legacy
`quarterTurns`, and require non-editable compatibility. Existing authored decks
without the new fields retain their previous conventional group behavior. Exact
source quarter turns retain the old representation where its integer centers
suffice; noncardinal and half-EMU centers use the new source representation.

## Source hierarchy

ECMA-376 Part 1 §§20.1.7.5–6 and Annex L.4.7.3–6 define the transform pipeline.
`ST_Angle` uses thousandths of an arcminute: 60,000 units per degree. The group
reference's 1/64,000 prose conflicts with that type definition; implementation uses
the type and the Annex's explicit 5,400,000-unit 90-degree example. Input angles
must be signed 32-bit integers. Source booleans are parsed exactly.

An individual mapping is `T = U⁻¹ R F U Tst`: child-coordinate scale/translation,
then center-relative flips, then clockwise rotation. Extents remain positive.

Annex L.4.7.4–5 deliberately specifies a different rule for a grouped leaf than
ordinary parent×child geometry. Horizontal/vertical scales and flips multiply on
the leaf's own axes, rotations add, and conventional nested matrices locate only
the final center. A 90-degree rotated 2×2 child inside a group scaled (2,1) has
matrix `[0,2,-1,0,3,-1]`; the conventional result `[0,1,-2,0,4,0]` has the same center
but different axes. Parsed groups therefore retain child order with identity paint
transforms, while their leaf nodes carry the complete source hierarchy transform.
This corrects the former parsed quarter-turn/nonuniform-group behavior. Legacy
authored groups continue to use conventional composition.

## Numeric and paint contract

`RenderTransform.sourceAffine` is the explicit `rational-affine-v1` variant. It
contains six reduced rational coefficients and six nonnegative absolute-error
bounds as canonical decimal numerator/denominator pairs, each bounded to 512 bits.
When present, the six legacy PPM fields must be identity. Consumers call the public
`renderTransformMatrix(transform, requestBudget)` helper; they must not discard the
rational variant and interpret its identity placeholders as the actual operation.
Exact results that fit the previous integer-PPM/integer-EMU representation retain
that representation.

Cardinal rotation, reflection, ratios, center translation, and rational algebra
are exact. Noncardinal sine/cosine values retain their exact binary64 rational
values with a disclosed conservative elementary-function allowance. This is a
bounded preview policy, **not a formal libm accuracy proof or Office parity claim**.
It does not round arbitrary rotation to integer PPM.

Composition propagates uncertainty including products of input errors. Paint
conversion adds the exact difference from the selected binary64 value. A
conservative depth-dependent arithmetic allowance then qualifies frame corners,
evaluated path/control envelopes, out-of-frame callouts, text regions, and supplied
positioned glyph control hulls. The compiler refuses uncertainty above 0.125 EMU
or coordinates beyond the caller's limit (at most 2⁴⁸−1 EMU). The paint worker also
retains its existing narrower output resource limits; limits are never clamped.

One `SourceAffineBudget` belongs to a complete compile or paint request: 1,000,000
conservative operation charges, 16 MiB encoded transport, and depth 64. Decode
rejects extra fields, sparse/accessor tuples, noncanonical or oversized fractions,
negative error bounds, and invalid depth. Paint stages and validates the complete
command stream before calling the supplied surface.

A measured Node 22.23.2 run on the development macOS host compiled 100/200 ordinary
noncardinal cubic callouts with out-of-frame control hulls in approximately
97/172 ms under the proposed 1m cap; the permanent 200-shape test passes, and a
larger request exercises aggregate refusal. Timings are local evidence, not a
performance guarantee.

## Text and editing authority

Text flip handling follows the pinned Apache POI implementation evidence in
[DrawTextShape.drawContent](https://github.com/apache/poi/blob/338882ac8898df5c13a7d15f533204c5dd8607d6/poi/src/main/java/org/apache/poi/sl/draw/DrawTextShape.java).
It XORs horizontal/vertical flips through the hierarchy and counter-reflects
horizontally about the source frame when total reflection parity is odd. Thus an H
flip does not mirror glyph outlines; V and H+V retain the corresponding 180-degree
text orientation. The counter-reflection precedes the existing text-only vertical
mapping. This is an implementation-backed preview policy, not a normative claim
about every Office application. Real-font worker tests verify H/V/H+V placement
with asymmetric insets and noncardinal text orientation.

Arbitrary/reflected text requires a supplied digest-bound glyph control-hull
provider. A source affine ancestor marks descendants preview-only, and the actual
mutation resolver independently rejects those targets from the ancestor chain;
removing a descendant diagnostic cannot grant mutation permission. Ordinary exact
group mutation behavior stays covered by the existing round-trip tests.

## Required remaining integration

- `upright=true` and separate nonzero body text rotation still require source and
  layout qualification. A text box with unsupported body semantics is refused;
  an AutoShape may retain geometry while explicitly omitting unsupported text.
- Newly complex groups containing table/chart graphic frames are preserved as one
  opaque subtree with a specific diagnostic. Their cell/axis text orientation and
  full hull qualification are required follow-up work. Existing exact groups stay
  supported; direct new affine graphic-frame inputs also refuse at compilation.
- The approximate `PptxFilePreview` still needs evaluated Q/C/A paths, relative fill
  modes, hierarchy placement, and explicit text limitations connected to its UI.
- Arbitrary custom geometry still needs sequential duplicate-guide reassignment;
  the preset catalog already preserves that behavior through alpha-renaming.

These remaining items prevent declaring the assigned complex-shape family complete.

## Browser SVG units

The existing native vector component renders in points (1 point = 12,700 EMU), while the worker transport remains in EMU. Every path/frame/clip/stroke length and matrix translation is divided by 12,700; matrix linear terms, arc flags/angles, stroke miter ratios, and image crop percentages remain unchanged. This uniform change of coordinate basis preserves nested glyph and shape placement and active clips. It avoids an observed Chrome clipping failure with large EMU user coordinates; no particular browser-internal cause is asserted. The actual WASM/worker browser proof checks all eight source shapes, normalized DOM dimensions against the EMU transport, and a separate visible-inside/hidden-outside clip control.
