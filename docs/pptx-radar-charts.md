# Source-bound radar charts: qualification work

The connected literal-only profile attaches qualified standard/filled radar to native extraction and opt-in rendering. Literal categories and numeric spellings remain bound to the original chart part, relationship and opaque fingerprint. A separate [embedded-workbook profile](PPTX-WORKBOOK-RADAR.md) resolves source references through exact saved XLSX cells; it never attaches those values as literal data.

## Source inventory

The ISO-derived [radarChart remarks](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.radarchart?view=openxml-3.0.1) identify radar style, series, varying colors, data labels and axis IDs. [Radar series](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.radarchartseries?view=openxml-3.0.1) carry index/order, optional title, shape properties, marker, point overrides, categories and values. Radar series have no line-series smooth element; a line-parser projection must not invent or ignore one.

The source style enum has `standard`, `marker` and `filled`. [Office requires an explicit radarStyle value](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/d8c4036e-c48c-48d8-9953-867045c8dd46). Its [255-series implementation limit](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/d1ada276-1367-4cc6-a9e1-7fbcca0d64b8) does not increase InjOffice's existing 16-series resource limit.

## Connected literal profile

* Dense 3–256 category/value pairs and 1–16 series. Preserve indexed point order and original series metadata; XML series sequence determines paint order in the retained references, while original numeric order metadata is retained.
* Standard closed series lines and filled closed polygons, with complete explicit RGB line/fill styles. Markers, point paint overrides, labels, grids, themes/defaults and extensions remain refused until their own semantics are qualified. This does not claim those remaining requirements are complete.
* One explicit category axis and one explicit linear numeric axis, linked by exact IDs. Reuse bounded decimal and source-axis primitives where their semantics match; do not project radial labels through Cartesian label layout.
* Literal values stay literal. The workbook parser returns category/value/title references only, with no cached data authority. The separate workbook profile uses the existing actual-XLSX engine, strict decoder and source-bound weak-set admission.
* Exact radial ratios precede any angular projection or final EMU quantization. Frame size, node and path budgets remain bounded. The first geometry profile requires every value within its explicit scale. Out-of-scale polygons remain refused instead of silently clamping data.

The Office automation owner supplied standard/filled/marker, 4/5-category, signed/zero, reversed-axis, XML-order permutation and axis-isolation cases. The qualified observations and retained normalization differences are recorded below.

## Shared handoff

The literal-radar record has Go/TypeScript/schema validation and strict exclusion from other literal families. `literalRadarPreview:true` enables public compilation; the supplied-font worker routes the existing source-chart toggle to this option. Default-off preserves opaque/image fallback. Existing source-frame normalization remains authoritative for own rotated/reflected and grouped frames; generated vector ink is qualified before plot clipping. The separate workbook opt-in retains its own inspection/resolution authority. Radial axis labels remain unsupported.

## Reference-qualified bounded rendering policy

Fifteen clean-container PowerPoint 16.112.4 reference exports are retained externally with source/export hashes and raw source/saved XML. Exact category/value/index/order/axis settings and source RGB values survive the first twelve cases. Full source trees do not: standard style becomes marker with explicit no-marker series retained; filled style removes explicit none markers; Office adds other formatting metadata. Those differences remain recorded and are not used to rewrite source authority.

The observed geometry starts category zero at the top and proceeds clockwise; reversed category orientation reverses that direction. Radius is `(value-min)/(max-min)` and reversed value orientation complements it, including signed values that lie within the explicit scale. Standard paths and filled polygons close to their first category. Filled series paint in XML sequence. The value axis owns the complete spoke set; category-axis line paint does not produce a separate path in the three colored/deleted-axis isolation fixtures. Standard spokes paint before data, filled spokes after data. Labels, grids and automatic axis defaults are outside this first profile.

`source-radial-plot-v1` fits a circle of radius half the smaller host plot extent at its center. It preserves these radial relationships rather than claiming Office plot-margin parity. Existing 4096-bit ChartRational arithmetic keeps radial ratios and final XY algebra exact. Cardinal directions are exact; other directions use an explicitly non-formal libm allowance matching the source-affine policy: `64*Number.EPSILON*(abs(angle)+1)`. At the 100-million-EMU frame limit, the angular allowance is below 0.00001 EMU. Coordinates round once to whole EMU, with at most 0.5 EMU algebraic rounding, and complete stroke hulls include a separate conservative one-EMU numerical outset. This is a qualification policy, not a proof of the JavaScript transcendental implementation. Independent 90-digit Decimal Taylor references test 3/5/7/256 directions at the maximum frame.

Each series has at most 257 commands; all 256 spokes fit exactly 512 commands. Up to sixteen series and one spoke node imply at most 4,624 commands. Generic 512-per-path limits do not change. The connected compiler must qualify complete ink hulls before plot clipping and retain the separate source/workbook admission boundary.


## Integration checks and remaining limits

Sixteen source packages cover strict/transitional XML, standard/filled styles,
plain/own-rotated/nonuniform grouped frames and reversed category/value axes.
Three differently sized colored series retain XML order `[2,0,1]`, source indices,
and raw numeric spellings. Actual Go and WASM extraction are compared across all
source/paint fields and fingerprints; their independently issued capability token
strings are deliberately excluded from comparison without modifying either deck.
The unchanged WASM deck reaches the supplied-font compiler and framed production
worker subprocess. Default-off, source immutability, full family exclusion and
out-of-scale/marker/extension/label/ref authority refusal are tested separately.

Public compiler tests retain the generic 512-command limit with 16 series × 256
categories plus the 512-command spoke node. Complete stroke hulls plus one EMU
numerical allowance are checked before plot clipping, including source-frame size
uncertainty. The extra allowance is attached only when literal radar is enabled
and rendered; default-off placeholders retain their existing bounds. Source axis
labels, markers, grids, themes/defaults, point overrides and outside-scale points
remain unavailable. Workbook radar uses its separate source-bound profile.

`scripts/smoke-pptx-literal-radar-browser.mjs` generates all sixteen sources and
uses the actual helper/worker and production SVG component. It compares each
series path's paint and geometry, scrolls the measured slide into view, captures
the slide region, and checks nondegenerate bounds and visible ink with same-rectangle
hide-one-series controls. Required changed pixels scale with the measured stroke
area or triangle area (five percent, minimum two pixels); exact RGB counts are
informational because subpixel stroke antialiasing can blend colors. Each control
restores its original visibility in a finally block. The
separate Linux qualification workflow builds fresh artifacts and preserves source,
DOM, preview, pixel counts and screenshots. Local Chrome startup remains blocked;
no browser success is claimed until that workflow passes and screenshots are
reviewed. The retained fifteen Office exports qualify the documented radial
relationships, not general PowerPoint plot layout equivalence.

## WASM packaging budget

The Go 1.23.0 darwin/arm64 toolchain (also the local toolchain) produced 7,873,433
bytes with SHA-256 `05590fb3832b43a759f99c7995ec3704ab0026a6f81fce13e151f4ed37202059`.
An initial redundant geometry-source validation pass produced 7,876,637 bytes;
removing that duplicate call saved 3,204 bytes. The parser and complete public
radar validation still enforce every source/reference/count/style/axis/range
condition, independently reviewed. This changes no trust boundary.
The previous 7.5 MiB artifact ceiling (7,864,320 bytes) was exceeded by 9,113 bytes.
The literal parser/schema adds executable code, so its build artifact ceiling is
now 7.75 MiB (8,126,464 bytes), providing 253,031 bytes of measured margin.
This is only the WASM distribution build gate; package/input, node/path, rational,
font and source-affine budgets and worker admission remain unchanged. Fresh
consumer/WASM checks and CI are still required; sizes can vary by toolchain.

### Ceiling raised to 8.25 MiB + 8 KiB

The 7.75 MiB ceiling (8,126,464 bytes) was consumed by three merged PPTX
features: connector presets (#214), picture clipping (#216) and the SmartArt
drawing fallback (#218). `main` at 3d8b9e32 measured 8,083,787 bytes on Go
1.23.0 darwin/arm64, leaving 42,677 bytes of margin, while the pending
table-style work (#220, rebased on that `main`) measured 8,150,024 bytes locally
and 8,155,244 bytes in CI, so it failed the gate. Pending autofit/placeholder
work (#221) adds another 184,747 bytes over its base (7,935,816 -> 8,120,563
bytes at c7b7bece), and a diagram-layout lane is in progress. The shared ceiling
in `go/pptxpatch/cmd/pptxnativewasm/max-bytes.txt` is therefore raised one
quarter-MiB step to 8.25 MiB + 8 KiB (8,658,944 bytes), matching the XLSX
convention. Projected `main` plus #220 plus #221 is about 8,334,771 bytes,
leaving roughly 324,000 bytes of measured margin for the diagram-layout lane;
`main` alone has 575,157 bytes of margin. The build script and CI gate now print
the ceiling read from that file instead of a hand-maintained figure.

### Ceiling raised to 8.5 MiB + 8 KiB

The SmartArt hierarchy layout evaluator (#227: bounded ECMA-376 §21.4 data
model, layout definition evaluation, composite/hierRoot/hierChild/sp/tx/conn
algorithms and style resolution) compiles to roughly 327,000 bytes of WASM.
`main` at 3953cc64 (with #220 and #226 merged) measured 8,150,024 bytes on Go
1.23.0 darwin/arm64; the #227 head measured 8,477,268 bytes locally (CI Linux
builds run about 5,000 bytes larger), leaving 181,676 bytes under the
8.25 MiB + 8 KiB ceiling (8,658,944 bytes). The pending autofit/placeholder
work (#221) adds 184,747 bytes, so `main` plus #227 plus #221 projects to about
8,662,000 bytes locally and 8,667,000 bytes in CI, a few hundred bytes over
that ceiling. As the largest consumer, #227 carries the next deliberate
quarter-MiB step: `go/pptxpatch/cmd/pptxnativewasm/max-bytes.txt` is raised to
8.5 MiB + 8 KiB (8,921,088 bytes), leaving 443,820 bytes of measured margin for
this head and roughly 254,000 bytes once #221 lands. The build script and CI
gate keep reading the ceiling from that file.
