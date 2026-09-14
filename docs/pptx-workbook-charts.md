# Embedded workbook chart source

The read-only `inspectChartWorkbooks(bytes)` PPTX client operation inspects
qualified bar/column, line and straight XY chart reference descriptors. It
returns exact local XLSX package bytes and source identities. It does not grant
chart editing authority, interpret SpreadsheetML, or use chart/formula caches.
Existing literal profiles remain separate.

The `chartWorkbooks` worker operation uses the existing request identity, queue,
abort, timeout and worker recovery protocol. Its Go binding calls
`InspectNativePPTXChartWorkbooks`. The response protocol is
`pptx-chart-workbook-inspection-v1`; every projected native chart appears either
as an admitted chart descriptor or an explicit omission. Source outside the
native chart projection remains covered by the original deck's preservation
diagnostics, rather than becoming an inspection capability.

The response binds outer package SHA-256 and revision; slide part and digest;
element/object identity; exact original `p:graphicFrame` XML digest; chart
relationship, part and digest; and workbook relationship, part, digest and byte
length. `frame_sha256` hashes exact source XML bytes, not a canonical JSON
descriptor. Workbook resources are deduplicated by exact package part. The
client validates source closure against a separately extracted deck from its
private source snapshot, checks resource hashes and canonical base64, rejects
duplicate JSON members and unsafe integers, and returns frozen admitted records.

Limits per inspection are 64 chart/omission records, 8 unique workbook resources,
8 MiB per workbook, 16 MiB aggregate workbook bytes and 32 MiB serialized response.
The caller's existing PPTX package limit still applies. Unknown fields and
unreported or duplicate chart/resource identities refuse.

`extractChartWorkbookReferences` accepts an explicitly supplied trusted XLSX
extractor callback. The callback receives a disposable copy of the selected
workbook bytes. A private snapshot, exact source SHA-256 and length, and frozen
reference descriptors survive asynchronous execution. Its bounded serialized
V2 result passes the public `@injoffice/sheets/browser` decoder before any cell
join. This imports the existing decoder, not another XLSX engine into PPTX WASM.
Validation cannot cryptographically prove an arbitrary callback performed the
claimed extraction; the caller explicitly selects that authority. Built-in
browser composition uses the actual separate XLSX WASM client, and server
composition must use the existing Go XLSX extractor.

References are direct worksheet-qualified one-dimensional A1 ranges. Exact
worksheet names, quoted names and escaped apostrophes are supported. Numeric
source values retain bounded decimal lexemes; categories use plain shared or
inline source strings. Referenced formulas (including usable-looking caches),
errors, dates, numeric text, rich strings and missing cells refuse. Unreferenced
formula cells do not invalidate a literal referenced range. Chart caches are
bounded and recorded as present, but their values and counts never authorize
workbook values. Resolution records carry `dataOrigin: embedded-workbook`, the
original formula, actual addresses and workbook revision/hash.

The first source profile requires explicit `plotVisOnly=false`. XLSX row/column
visibility values are retained as dimension-record metadata, without claiming
their hidden attributes were explicitly written and without applying a guessed
filter. `dispBlanksAs` is retained; missing cells still refuse rather than being
silently interpreted as gaps, zeros or spans. Resolution caps 64 references,
256 values per reference, 16,384 total values, 2 million source records,
32,768 category UTF-16 units and 4,096 returned source diagnostics.

Current diagnostic qualification is deliberately conservative. Referenced
opaque cells/formula ranges and unknown workbook/worksheet/visibility source
features refuse; harmless style/protection/decoration diagnostics remain visible
as source provenance. XLSX now emits distinct `WORKBOOK_VIEW_METADATA` and
`WORKSHEET_DIMENSION_METADATA` records for closed qualified view/used-range
subtrees. These records remain preserved; adjacent generic unknown records
remain separate and still refuse. Qualification accepts only known, typed,
unqualified attributes, native namespaces and XML whitespace. Extension/MC/
foreign content is not promoted. Ordinary workbook metadata is never stripped.

`resolveNativePptxWorkbookCharts(inspection, extractor)` extracts each unique
resource once and returns separately admitted charts and source-scoped refusals.
Its aggregate budget is 16,384 referenced points and 2 million scanned records.
The public `createResolvedWorkbookChart` factory accepts only admitted inspection
and value records; plain JSON copies cannot bypass either decoder. Workbook data
uses separate `workbook-bar-v1`, `workbook-line-v1` and `workbook-scatter-v1`
profiles. Original literal records and their source guards are unchanged.

Pass the admitted records as `workbookChartsPreview` to
`compileNativePptxSlide`; omission retains opaque image/placeholder fallback.
Compilation binds package revision and the exact slide, frame and chart source
fingerprints again. It uses shared exact decimal geometry, source paint, integer
endpoint rounding and plot stroke clipping. Axis labels additionally require
`chartAxisLabelsPreview` and exact supplied fonts under the existing bounded
Latin/Common/Inherited label policy. Plot margins and placement remain disclosed
host layout, not PowerPoint layout equivalence.

The playground helper mode `charts=source-workbook` is separately selected from
`charts=source-literal`. Go runs both existing extraction engines, then the paint
worker validates their JSON and source bindings through the same public adapters.
No pre-resolved geometry crosses that trust boundary. The aggregate serialized
UTF-8 workbook payload is capped at 8 MiB inside the existing 16 MiB worker
request/response limits. The strict `workbook_chart_preview` response flag must
match the request. Operator font configuration and upload consent are unchanged.

The metadata classification also exposes an existing cell-writer behavior: adding
its exact recalculation `calcPr` XML can introduce a previously deduplicated generic
workbook record. Only cell transactions whose produced workbook XML is byte-for-
byte the existing `workbookWithFullRecalculation` result may add that one derived
record. All old inventory records and the default strict comparator remain intact;
`calcPr` has not been classified as chart-neutral.

Formula recalculation, defined names/noncontiguous ranges, visibility filtering,
blank semantics, inherited styles, automatic axes, additional chart families and
full chart typography remain required follow-up work. This milestone connects
qualified saved workbook values to native vectors; it does not complete all native
chart support.
