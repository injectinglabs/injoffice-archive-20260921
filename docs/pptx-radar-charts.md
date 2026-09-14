# Source-bound radar charts: qualification work

This dedicated preparation is not yet attached to native extraction or public rendering. The intended connected milestone admits literal categories/numbers and separately resolved embedded-workbook references, preserving raw source values and package identity. Existing chart profiles remain unchanged.

## Source inventory

The ISO-derived [radarChart remarks](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.radarchart?view=openxml-3.0.1) identify radar style, series, varying colors, data labels and axis IDs. [Radar series](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.radarchartseries?view=openxml-3.0.1) carry index/order, optional title, shape properties, marker, point overrides, categories and values. Radar series have no line-series smooth element; a line-parser projection must not invent or ignore one.

The source style enum has `standard`, `marker` and `filled`. [Office requires an explicit radarStyle value](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/d8c4036e-c48c-48d8-9953-867045c8dd46). Its [255-series implementation limit](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/d1ada276-1367-4cc6-a9e1-7fbcca0d64b8) does not increase InjOffice's existing 16-series resource limit.

## Proposed first connected profile

* Dense 3–256 category/value pairs and 1–16 series. Preserve indexed point order and original series metadata; XML series sequence determines paint order in the retained references, while original numeric order metadata is retained.
* Standard closed series lines and filled closed polygons, with complete explicit RGB line/fill styles. Markers, point paint overrides, labels, grids, themes/defaults and extensions remain refused until their own semantics are qualified. This does not claim those remaining requirements are complete.
* One explicit category axis and one explicit linear numeric axis, linked by exact IDs. Reuse bounded decimal and source-axis primitives where their semantics match; do not project radial labels through Cartesian label layout.
* Literal values stay literal. The workbook parser returns category/value/title references only, with no cached data authority. Resolution must later use the existing actual-XLSX engine, strict decoder and source-bound weak-set admission.
* Exact radial ratios precede any angular projection or final EMU quantization. Frame size, node and path budgets remain bounded. The first geometry profile requires every value within its explicit scale. Out-of-scale polygons remain refused instead of silently clamping data.

The Office automation owner supplied standard/filled/marker, 4/5-category, signed/zero, reversed-axis, XML-order permutation and axis-isolation cases. The qualified observations and retained normalization differences are recorded below.

## Shared handoff

After stacked-chart integration releases the shared paths, add separate literal-radar and workbook-radar records and strict family exclusion, source extractor and workbook-inspection hooks, native/Go/schema validation, opt-in compiler and worker/UI routing. Retain graph-frame transform refusals and source ownership. Actual dual-WASM and recorded-paint/browser evidence must accompany a connected PR; helper-only publication is not a milestone.

## Reference-qualified bounded rendering policy

Fifteen clean-container PowerPoint 16.112.4 reference exports are retained externally with source/export hashes and raw source/saved XML. Exact category/value/index/order/axis settings and source RGB values survive the first twelve cases. Full source trees do not: standard style becomes marker with explicit no-marker series retained; filled style removes explicit none markers; Office adds other formatting metadata. Those differences remain recorded and are not used to rewrite source authority.

The observed geometry starts category zero at the top and proceeds clockwise; reversed category orientation reverses that direction. Radius is `(value-min)/(max-min)` and reversed value orientation complements it, including signed values that lie within the explicit scale. Standard paths and filled polygons close to their first category. Filled series paint in XML sequence. The value axis owns the complete spoke set; category-axis line paint does not produce a separate path in the three colored/deleted-axis isolation fixtures. Standard spokes paint before data, filled spokes after data. Labels, grids and automatic axis defaults are outside this first profile.

`source-radial-plot-v1` fits a circle of radius half the smaller host plot extent at its center. It preserves these radial relationships rather than claiming Office plot-margin parity. Existing 4096-bit ChartRational arithmetic keeps radial ratios and final XY algebra exact. Cardinal directions are exact; other directions use an explicitly non-formal libm allowance matching the source-affine policy: `64*Number.EPSILON*(abs(angle)+1)`. At the 100-million-EMU frame limit, the angular allowance is below 0.00001 EMU. Coordinates round once to whole EMU, with at most 0.5 EMU algebraic rounding, and complete stroke hulls include a separate conservative one-EMU numerical outset. This is a qualification policy, not a proof of the JavaScript transcendental implementation. Independent 90-digit Decimal Taylor references test 3/5/7/256 directions at the maximum frame.

Each series has at most 257 commands; all 256 spokes fit exactly 512 commands. Up to sixteen series and one spoke node imply at most 4,624 commands. Generic 512-per-path limits do not change. The connected compiler must qualify complete ink hulls before plot clipping and retain the separate source/workbook admission boundary.
