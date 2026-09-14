# Source literal bar and column previews

`NativeOpaqueChart.literalBar` carries a read-only `literal-bar-v1` profile extracted from qualified `c:barChart` source. `createNativeLiteralBarPaths` draws clustered bars or columns from that record. `compileNativePptxSlide` requires `literalBarPreview: true`; the playground has an explicit source chart preview checkbox. Original chart ownership and packaged preview fallback remain preserved.

This profile adds multiple series, signed/fractional/scientific literal values, indexed category strings, explicit linear value scales, category/value orientation, individual point colors, and explicit axis visibility. It does not reproduce Office plot-area layout. The preview fits the plot to the full supplied frame. The source table shown by the playground is host UI, separate from slide labels.

Qualification requires:

- Explicit clustered grouping, bar/column direction, `varyColors=0`, gap width, and zero overlap; no labels, legend, 3D, trendlines, errors, pictures, effects, extension clauses or other families in the plot.
- One to 16 series, each with 1 to 256 dense indexed literal values and identically indexed literal categories. Source point elements may arrive out of order. Series IDs are unique; explicit series orders form a complete contiguous metadata permutation; XML series sequence determines host cluster slots and painter order. Duplicate category labels retain separate point identities.
- `numLit` with `General` format and `strLit` categories. Encoded `_xHHHH_` category/title escapes, references, caches, missing/error/sparse data, automatic category inference, external workbooks and formulas are not promoted. Source number spelling is retained in string-valued JSON fields; the native contract still uses integer-only JSON numbers.
- Complete local solid RGB/no-line series paint and/or complete indexed point overrides. `invertIfNegative=0` is explicit. Theme/default color inference, partial style cascade and inverted negative paint are not implemented.
- Reciprocal category/value axis IDs, explicit orientation and visibility, explicit finite min/max including zero, category crossing at the minimum category boundary, value-axis `crossesAt` exactly zero, and `crossBetween=between`. Automatic/log/date/secondary scales, nonzero crossings and midpoint crossing are refused.
- Deleted axes omit unused style clauses in this bounded profile. Visible axes explicitly suppress both tick kinds and labels, and supply complete RGB solid flat-ended line paint. Missing flags or paints never imply invisible axes.

Numbers are bounded to 128 ASCII bytes, 32 mantissa digits and an explicit exponent from -100 to 100. Decimal coefficients and coordinate ratios use exact integer arithmetic. Categories are bounded to 32,768 UTF-16 units in total; series titles to 1,024 units each. The maximum is 4,096 source points plus two axis paths, within existing renderer node budgets; each bar uses five path commands. No render budget was increased for this profile.

Gap width is a percentage of **one bar width**, not a percentage of the whole category band. With `s` series and zero overlap, the band is `(s + gapWidth / 100)` bar widths. Half of the gap lies at either side of the cluster. Category and series coordinates are computed as exact ratios and rounded once to integer EMU. Strict OOXML uses percent-suffixed gap/overlap values; Transitional also permits the legacy integer spelling.

Bars extend from zero and clip at the explicit min/max. Reversal applies before rounding. Zero or sub-EMU heights retain their source point identity but can have no visible area. Frames too narrow to retain one EMU per bar use the preserved preview fallback with a diagnostic. Source axis strokes can extend half their width beyond the fitted plot boundary.

This is the first Cartesian milestone, not completion of chart support. Source typography for ticks/categories/titles, automatic scales, stacking, percentage stacking, line/area/scatter/bubble/radar/stock/surface families, workbook provenance and wider paint semantics remain tracked in the chart completion matrix.

Primary basis: ECMA-376 Part1 §21.2.2.75 (gap width), §21.2.2.131 (overlap), §21.2.2.34 (crossing value), and CT_BarSer/CT_BarChart/axis ordered content models. Microsoft mirrors the [gap width definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.gapwidth?view=openxml-3.0.1) and [axis-local crossing units](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.crossesat?view=openxml-3.0.1). Host frame fitting and EMU quantization are implementation policies, not a claim of PowerPoint parity.
