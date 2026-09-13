# DrawingML source transform qualification

This document describes dedicated implementation preparation. Source extraction,
native contract and paint integration are pending; the helpers alone do not expand
the currently supported public preview surface.

The retained ECMA-376 Part 1 source defines shape/group transforms in §§20.1.7.5–6
and explains the transformation pipeline mathematically in Annex L.4.7.3–6.
`ST_Angle` uses thousandths of an arcminute: 60,000 units per degree. The grouped
transform reference's 1/64,000 prose conflicts with that definition; the implementation
uses the type definition and the Annex's explicit 5,400,000-unit 90-degree example.
Source angles are signed 32-bit integers, normalized to [0,21,600,000). Flips are
strict DrawingML booleans, defaulting to false.

For an individual transformation, scale/translation maps the child coordinate box
to the group's frame; center-relative flips then precede clockwise rotation:
`T = U⁻¹ R F U Tst`. Positive shape extents and child extents remain required.
Zero-extent degeneracies need separate geometry qualification.

**Nested source shapes do not use conventional affine composition for their whole
geometry.** Annex L.4.7.4–5 expressly makes the leaf's own-axis horizontal/vertical
scales and flips the products along the hierarchy, and its rotation the sum of the
angles. Conventional nested transforms locate only the leaf's center. The dedicated
`sourceHierarchyAffine` implements that distinction. A 90-degree rotated 2×2 child
inside a group scaled (2,1) has matrix `[0,2,-1,0,3,-1]`, not conventional
`[0,1,-2,0,4,0]`; both have center (2,1), but scale along different axes.
The existing parsed quarter-turn/nonuniform-group path currently uses conventional
composition. Correcting it requires a separately reviewed compatibility/golden change;
the dedicated preparation does not silently alter that behavior.

The proposed precise transport contains six reduced rational coefficients and six
nonnegative absolute-error bounds, all as canonical decimal numerator/denominator
pairs with a 512-bit bound. Cardinal rotation, reflection, group ratios, center
translations and algebraic cancellation are exact. Noncardinal sine/cosine values
are represented by their exact binary64 rationals with a disclosed conservative
elementary-function allowance. This is a bounded preview policy, **not a formal
proof of JavaScript libm accuracy or Office visual parity**. It avoids rounding
arbitrary rotations to the old transport's integer parts-per-million coefficients.

Rational composition propagates coefficient uncertainty exactly, including the
product of input uncertainties. Converting coefficients for SVG/paint adds the exact
difference between the rational value and the selected binary64 value. Point
qualification adds a conservative depth-dependent paint arithmetic allowance,
refusing an accumulated bound above 0.125 EMU or a world envelope outside the caller's
coordinate limit (at most 2⁴⁸−1 EMU). Integration must qualify geometry/control hulls,
out-of-frame callouts and text placement, not only frame corners. Existing exact PPM
transport paths remain available independently.

One shared `SourceAffineBudget` must belong to the entire compile request. It limits
the aggregate conservative operation charges to 100,000, encoded transport to
16 MiB, and hierarchy depth to 64. Isolated helper defaults are for isolated calls;
integration must explicitly pass the request-wide instance throughout. Decode
rejects unknown fields, sparse/accessor tuples, noncanonical/oversized fractions,
negative errors and invalid depth. No JSON BigInt extension is required.

Text cannot simply inherit every shape/group flip. Source `upright`, body rotation,
text placement and readable glyph orientation require their own qualification.
Existing extraction refuses unsupported upright/body properties. That text integration
and the existing file-preview UI are required remaining parts of this milestone.
