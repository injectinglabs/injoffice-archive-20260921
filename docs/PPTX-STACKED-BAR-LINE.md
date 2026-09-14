# Stacked Cartesian bar and line preparation

This dedicated source/geometry layer is not yet attached to the native deck,
compiler or browser. Existing literal and workbook profiles remain unchanged.
A connected source/WASM/compiler/browser integration is required before shipping.
Workbook-backed stacked charts remain a required subsequent connected profile;
this literal preparation does not authorize reference caches or close that row.

The bounded profiles use complete aligned literal categories and decimal values,
at most 16 series and 256 categories. Bar series retain explicit base/point RGB
fills, no outline and `invertIfNegative=0`; bars require explicit full overlap
100 and gap width 0–500. Lines retain explicit straight RGB strokes, no markers,
no smoothing and the existing stroke limits. Both require reciprocal explicit
linear axes containing zero, supplied source orientation and source axis paint.
Referenced caches do not gain literal authority.

## Signed accumulation and exact source order

PowerPoint 16.112.4 (build 16.112.26090911) was used for isolated local reference
exports. The first 24 cases cover ordinary and percentage bar/line/area charts
with ordered values `[4,-2,3,-1]`, all-negative `[-2,-3]`, cancelling `[2,-2]`
and zero `[0,0]`. All original chart data and axis semantics survive Office save.
The additional ten cases cover horizontal bars, reversed axes, XML series order,
category-dependent totals and area zero crossings. Raw PPTX/PDF, generators,
Office-saved parts and Poppler vector decompositions are retained externally in
`pptx-stacked-office-20260913`; the two reference proof JSON files pin hashes.

The canonical 24-case source SHA-256 is
`dd895518d88b50acff4c18700399845a3802d851aa2501abde7a5db364416f4a`;
its actual Office PDF is
`df880856a7f377539414ea5a84b91e4e50c4b28600421902a9216e0da8b76632`.
The ten-case extension source is
`81df388a09b678319b342ceea226d9480326138cde9cf881bd28140bc9639f6f`;
its actual Office PDF is
`44b40bc843ff7e5ee3e78a2ec4e792e4be4b242a04fef47c8d723690e8122661`.

* Bars accumulate positive and negative inputs separately from zero.
* Lines use one algebraic running total, retaining each negative contribution.
* Both percentage families divide by the category's sum of absolute inputs.
  The denominator is shared across signs; a zero denominator produces zero
  boundaries. No decimal rounding or sign substitution occurs.
* Accumulation and paint follow the XML `<ser>` sequence. Original `c:idx` and
  `c:order` are retained unchanged. `c:order` must form a unique contiguous
  permutation, but it need not match the array sequence. The reversed XML
  reference proves Office follows that sequence while preserving differing
  `c:order` metadata on save.

The new geometry uses exact rational cumulative boundaries and category
coordinates. It clips before one integer-EMU quantization. Horizontal and
reversed axes change projection, not accumulation. Zero-width clipped bars have
no filled geometry; zero line series remain baseline lines when two or more
categories exist. Singleton and fully clipped line paths are empty and must be
filtered at compiler attachment. A clipped gap never gains an artificial line.
Each bar is five commands; each line has at most 510 commands. Existing generic
512-command, transport, coordinate and rational budgets remain unchanged.

Legacy helpers are reused only after a strict local preflight rejects sparse
arrays, trailing-newline RGB colors, negative-zero numeric identities and unknown
record fields. Temporary zero-value/normalized-order scaffolds are private
copies used for existing axis/frame validation; they never replace authoritative
source data, IDs, order metadata, accumulation or paint. Native label validation
retains the existing full supplied-font axis-label rules.

## Reference boundaries and required follow-ups

These observations qualify this source profile; they do not claim universal
Office layout parity. Geometry still fits the full host plot frame and follows
the existing explicit host axis-label layout policy. In the reversed reference,
Office normalizes category-axis `axPos=b` to `t` on save; the raw difference is
recorded instead of being reported as exact axis preservation.

The reference also establishes two required follow-ups outside these dedicated
files: existing chart families sort numeric `c:order`, and the current area
preview assumes a zero baseline. Actual area first-series fill follows the
authored category-axis crossing (`min` versus `autoZero`), including zero-input
behavior. Neither discrepancy is silently corrected here or treated as closed.

ECMA-376 Part 1 sections 21.2.3.4 and 21.2.3.17 describe the grouping modes;
the mixed-sign and XML-order observations above come from the retained actual
application reference. Microsoft documents the local PDF export mechanism in
[Save PowerPoint presentations as PDF files](https://support.microsoft.com/en-us/powerpoint/training/save-powerpoint-presentations-as-pdf-files).
