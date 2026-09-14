# Stacked charts from saved embedded workbooks

Qualified stacked and percentage bars, columns and lines can use saved embedded
workbook values through the existing explicit workbook preview opt-in. Source
chart caches remain evidence only. Literal previews and worksheet authoring
permissions do not change.

The PPTX inspector records the exact source grouping, full bar overlap 100,
category/value references and source paint/axes. New stacked descriptors preserve
XML series sequence and original unique contiguous `c:order` metadata. Ordinary
chart descriptors keep their previous shape and behavior. The closed decoder
requires grouping/overlap only in their qualified families; missing fields do not
imply a new stacked profile.

Resolution binds each formula/range to the unchanged embedded workbook part,
relationship, SHA-256 and saved revision. Actual XLSX extraction supplies cells;
`numCache`/`strCache` values never enter geometry. The admitted result uses distinct
`workbook-stacked-bar-v1` or `workbook-stacked-line-v1` profiles with
`dataOrigin: embedded-workbook`. Its values and source metadata are frozen. Private
geometry projection removes the outer authority fields without manufacturing a
literal profile or rewriting the source object.

Signed accumulation follows the same qualified policy as
[PPTX-STACKED-BAR-LINE.md](PPTX-STACKED-BAR-LINE.md): separate positive/negative totals
for bars, algebraic cumulative tops for lines, and each category's sum of absolute
values as the percentage denominator. Zero totals produce zero boundaries. Exact
rational accumulation precedes clipping and integer-EMU rounding. XML series
sequence determines accumulation and paint while original IDs/order are retained.
Horizontal and reversed source axes remain projections of the same data.

Supplied-font axis labels use the existing measured host plot layout. Unsupported
cells, formulas, metadata or fonts retain identified refusals; a refused chart
cannot recover cached data authority. Another valid chart in the same workbook
can still render. Source presentations and embedded workbooks remain unchanged.

## Limits and evidence

The existing 8 MiB per-workbook, 16 MiB aggregate resource and 8 MiB worker UTF-8 payload
limits remain independently enforced. Existing chart/reference/point/scan,
4096-bit rational, generic 512-command path and affine bounds are unchanged.
The built PPTX WASM measured 7,804,789 bytes under the existing 7,864,320-byte ceiling
with Go 1.23.0. No dependency or package budget was increased.

Thirty-six source fixtures cover columns/horizontal bars/lines, stacked/percentage
modes, mixed signs, all-negative inputs, cancellation, zeros, XML permutations and
complete supplied-font labels. They retain ordinary workbook metadata from the
existing corpus fixture while explicitly authoring the test worksheet cells.
Every chart contains contradictory 999 cache values. Actual browser PPTX and XLSX
WASM workers admit the saved values, retain exact source bytes and preserve the
source series sequence.

Six labeled source decks pass the actual font worker and the production
`NativePptxVector` React component; all A/B/C labels are visible. The actual Go
server's PPTX/XLSX extraction path feeds the same six profiles into the worker,
with substituted workbook hashes refused. External evidence lives in
`pptx-workbook-stacked-charts-20260913`, including fixture inputs, source-bound
engine outputs, font hashes, request logs and screenshots. The local test overlay
used for server evidence disables vet only because Go vet cannot open its virtual
new test path; production source was not edited for that probe.

This extension does not silently change the existing ordinary-family `c:order`
policy or the area chart's authored-crossing baseline. Those reference-backed
follow-ups remain separate required work.
