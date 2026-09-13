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

Limits per inspection are64 chart/omission records,8 unique workbook resources,
8MiB per workbook,16MiB aggregate workbook bytes and32MiB serialized response.
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
silently interpreted as gaps, zeros or spans. Resolution caps64 references,
256 values per reference,16,384 total values,2million source records,
32,768 category UTF-16 units and4,096 returned source diagnostics.

Current diagnostic qualification is deliberately conservative. Referenced
opaque cells/formula ranges and unknown workbook/worksheet/visibility source
features refuse; harmless style/protection/decoration diagnostics remain visible
as source provenance. Generic XLSX diagnostics currently combine ordinary
metadata with unqualified semantics. Ordinary workbook admission is pending an
upstream authority classification; production code never strips source metadata
to obtain admission.

Native resolved chart rendering, server composition and ordinary source
qualification are the remaining integration work in this milestone. Formula
recalculation, defined names/noncontiguous ranges, visibility filtering, blank
semantics, inherited styles, automatic axes, additional chart families and full
chart typography remain separate required work. Inspection alone does not
complete native chart support.
