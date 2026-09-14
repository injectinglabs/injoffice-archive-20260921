# Stacked Cartesian and area preparation

The dedicated core computes exact boundaries for complete aligned series and
qualifies literal area source data. It is not yet attached to the public native
contract, WASM extraction, compiler, or preview. No new chart option is available
from this preparation alone.

`standard` area uses an independent zero baseline and admits signed source
decimals. `stacked` accumulates nonnegative values in ascending authored series
`order`, separately for each category. `percentStacked` divides cumulative
boundaries by the exact total for that category, on the unit scale where 1 is
100%. A zero total yields only zero boundaries and no filled geometry. Explicit
source axes remain independent; no scale is inferred from an empty result.

Source decimal lexemes are retained. The arithmetic uses bounded exact integers
and rationals; rounding occurs only when clipped coordinates become integer EMU.
The core rejects negative stacked input pending family-specific evidence for
mixed-sign accumulation and percent denominators. This is a remaining scope row,
not a claim that mixed-sign stacking is complete.

Area source qualification requires one area family, dense matching literal
category/value lists, unique source indices and contiguous authored order,
explicit uniform RGB fill with no outline, and the existing qualified explicit
Cartesian axes. It refuses point-specific formatting, inferred theme paint,
missing data, caches substituted for source values, extra chart families, and
unrecognized or duplicate source clauses. Source bytes are not modified.

For geometry, categories lie at their between-category centers. Each interval
forms a band between its top and bottom boundaries. A signed standard interval
crossing zero is split exactly at that crossing to avoid a self-intersecting
polygon. Convex pieces are clipped to the plot rectangle before quantization.
Opposite shared edges cancel using exact endpoints and original face links;
point-only contacts remain separate rings. Collinear interval vertices are
removed, and zero-area results are omitted. All rings for one series must be filled as a
single compound path without outlines, so artificial interval borders are not
painted and separate antialiasing does not introduce seams. The private core
enforces at most 1,536 generated area commands per series for at most 256
categories. The generic custom-path limit remains unchanged. Standard area
series can overlap: the integration must state its source-order paint policy
separately; authored series order alone does not establish Office's paint order.

The primary audit uses ECMA-376 Part 1 (2016), §21.2.3.4 (bar grouping) and
§21.2.3.17 (grouping). These define stack placement and 100% scaling but do not
specify a mixed-sign algorithm. Microsoft's [Area record](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/43a435d7-33b8-41d5-b6e6-bc0bb9fb9f4a)
also describes same-category totals; it does not settle the OOXML sign cases.
Microsoft's [blank-cell support notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b5c5c694-21d9-437c-9a4a-21e0e843eed8)
distinguish standard area gaps from stacked area zero behavior. Dense source
qualification currently avoids choosing an unstated missing-data policy.
