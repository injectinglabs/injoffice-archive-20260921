# Source literal area previews

The native source extractor and WASM attach a read-only `literal-area-v1` record
when its complete source profile qualifies. The renderer exposes
`createNativeLiteralAreaPaths` and the default-off `literalAreaPreview` compiler
option. The playground literal panel and explicit source-literal measured
preview use the same complete-series geometry. Labels additionally require
`chartAxisLabelsPreview` and the existing supplied-font outline layout. Opaque
chart ownership and editing permissions remain unchanged.

`standard` area uses an independent zero baseline and admits signed source
decimals. `stacked` accumulates nonnegative values in XML series
sequence, separately for each category. `percentStacked` divides cumulative
boundaries by the exact total for that category, on the unit scale where 1 is
100%. A zero total yields only zero boundaries and no filled geometry. Explicit
source axes remain independent; no scale is inferred from an empty result.

Source decimal lexemes are retained. The arithmetic uses bounded exact integers
and rationals; rounding occurs only when clipped coordinates become integer EMU.
The core rejects negative stacked input pending family-specific evidence for
mixed-sign accumulation and percent denominators. This is a remaining scope row,
not a claim that mixed-sign stacking is complete.

Area source qualification requires one area family, dense matching literal
category/value lists, unique source indices and a complete contiguous authored order permutation,
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
painted and separate antialiasing does not introduce seams. The source-qualified renderer
enforces at most 1,536 generated area commands per series for at most 256
categories. The generic custom-path limit remains unchanged. Standard area
series can overlap: this preview paints in XML series sequence as
an explicit host policy. Authored order alone does not establish Office paint
order or visual parity. Stacked bar/line families and mixed-sign stacking remain
separate completion rows; this area option does not claim to complete them.

The primary audit uses ECMA-376 Part 1 (2016), §21.2.3.4 (bar grouping) and
§21.2.3.17 (grouping). These define stack placement and 100% scaling but do not
specify a mixed-sign algorithm. Microsoft's [Area record](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/43a435d7-33b8-41d5-b6e6-bc0bb9fb9f4a)
also describes same-category totals; it does not settle the OOXML sign cases.
Microsoft's [blank-cell support notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/b5c5c694-21d9-437c-9a4a-21e0e843eed8)
distinguish standard area gaps from stacked area zero behavior. Dense source
qualification currently avoids choosing an unstated missing-data policy.

## Resource limits and verification

A generated area path has at most 1,536 commands for at most 256 categories.
The bound covers original vertices, exact clipping/zero-crossing copies and
ring closes: signed standard bands need at most `6n−2` commands, and the
nonnegative stacked profile at most `5n`, where `n≤256`. Runtime admission
checks the actual result as well as these source limits.
The generic arbitrary custom-path ceiling remains 512. The worker accepts
200,000 SVG characters per path and 8 million total, with 16 MiB framed and
browser response limits; none of these limits changed. The 256-category signed
clipping fixture produces one 1,280-command compound fill and passes the real
worker decoder. Zero-area bands produce no fill while explicit axes remain.

The PPTX-only WASM ceiling is 7.5 MiB (7,864,320 bytes), enforced by the same
`go/pptxpatch/cmd/pptxnativewasm/max-bytes.txt` in the local build and CI. On
Go 1.23.0 darwin/arm64, the integrated workbook/guide baseline is 7,594,533 raw
bytes (2,024,934 gzip-9 bytes); the connected area artifact is 7,650,837 raw bytes (2,034,994 gzip-9
bytes, deterministic zero timestamp). Removing unused extraction-time rational
band computation saved 11,603 bytes without changing admission: the source
already qualifies counts, indices, dense alignment and decimal lexemes, then
checks a complete contiguous series order permutation and the supported nonnegative stacking domain.
Exact band computation remains tested and the renderer retains exact rational
geometry. No validation was removed to fit a package budget. DOCX/XLSX ceilings
are unchanged. Artifact sizes depend on the Go toolchain.

`node scripts/smoke-pptx-area-browser.mjs` generates actual source fixtures,
loads them through the browser WASM extractor, and exercises explicit local
server/worker rendering with installed DejaVu font bytes. It verifies source
hashes, default-off consent, exact numeric spellings, zero axes, unsupported
negative-stack fallback, supplied-font labels and the complete large compound
path. Screenshots and source artifacts are written outside the repository.
