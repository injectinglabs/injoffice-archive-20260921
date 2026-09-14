# Read-only source-style preview V1

`xlsxpatch.PreviewNativeSourceStylesV1(originalBytes)` returns a separate
`injoffice.xlsx.source-style-preview` envelope (`version: 1`, `read_only: true`,
`fidelity: approximate`). It is a Go API; browser integration is a follow-on.
It never constructs a `NativeWorkbookV2`, revision, mutation capability, or
rewritten ZIP. Strict extraction and mutation retain their existing refusal.

The source-backed case is the unchanged office2pdf expense report with child
fill and number-format IDs that differ from their parent while the corresponding
apply flags are absent. This preview selects those direct IDs for display and
records every conflict, including parent/direct IDs and the absent flag name.
An explicit false conflict is refused. Font, border, and alignment inheritance
remain subject to the strict rules.

## Source evidence and supported content

The envelope includes the package SHA-256 and resolved workbook/styles parts.
Each projected style includes its source and parent IDs/hashes, selected
font/fill/border IDs/hashes, and the selected number-format ID/code. Cell values
remain source text/lexicals; formula text and saved results are labeled caches
and are never recalculated. Consumers must disclose possible stale caches.

The bounded profile accepts one visible Transitional worksheet, at most 128
rows, 32 columns, 4096 stored cells, 128 merges, and 128 records per relevant
style table. Cell text, formula, format and aggregate text limits also apply.
It projects literal values, ordinary formulas with saved results, explicit
stored heights/widths, merges without covered content, simple fonts, direct
opaque RGB solid fills, and direct RGB thin borders. Unresolved font color is
explicitly reported with a black display fallback. Number-format strings are
source records, not a promise that a downstream formatter implements them.

OPC routes, XML, content types, style references, coordinates, and merges pass
the existing validated readers. A separate closed worksheet/style projection
refuses unsupported structural/visual content, rich strings, formula groups,
hidden/outline bands, column/row styles, drawings, and unqualified metadata.
The entire preview fails when any required record falls outside the profile.

## Display limits

Host font matching, column-width conversion, wrapping, and border metrics are
approximate. Print settings, view settings, calculation, and external updates
are not applied. This source grid is not a printed-page preview or evidence of
Excel-calibrated pagination. Original bytes remain the only file authority.
A consumer must keep this envelope separate from native editing/export state.

The original expense source SHA-256 is
`3403db276febaab11fedea12efc734af1a99fe0f25f2fc306c2093dd354ca283`.
The public Go API reports seven conflicts, the navy `#1E2761` fill, three merges,
and saved SUMIF results `2229.7` and `249.5`; strict extraction still refuses.
Tests cover the separate authority boundary and malformed/unsupported sources.
No browser visual or Excel print-calibration claim follows from this API test.
