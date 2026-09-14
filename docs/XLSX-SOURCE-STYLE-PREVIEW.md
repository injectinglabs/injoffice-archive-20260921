# Read-only source-style preview V1

`xlsxpatch.PreviewNativeSourceStylesV1(originalBytes)` returns a separate
`injoffice.xlsx.source-style-preview` envelope (`version: 1`, `read_only: true`,
`fidelity: approximate`). The Go API and the separate browser WASM client share this envelope.
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

## Browser recovery and verification

The playground offers an explicit browser-local source preview after strict
extraction refuses. A newly opened source clears the old editing session before
extraction. Recovery uses a dedicated worker that accepts only read-only
inspection messages, and a separate closed JSON decoder with source-hash,
style/cell/conflict/merge joins. Results are recursively frozen. Cancel,
replacement and unmount terminate the worker and invalidate late results.

The HTML grid displays source styles and merged anchors once. Source values,
formula text, conflict IDs and display fallbacks remain inspectable. An exact
source-declared euro suffix is supported for browser text without changing the
strict glyph formatter. All other unsupported formats fall back with a visible
warning. Host typography/geometry remains approximate; there is no editing,
repair/export or print-calibration authority.

Synthetic CI browser qualification opens an editable source, replaces it with
the conflicting source, verifies old edit controls disappear, renders through
the actual dedicated WASM worker, checks merge/color/cache output and no upload,
and delivers late replies after cancellation/replacement. Screenshot artifacts
support visual review. External benchmark files are not loaded in CI.

## Conditional source evidence V2

`xlsxpatch.PreviewNativeSourceStylesV2(originalBytes)` returns version 2 of the
same read-only protocol. `grid` is the V1 base-style envelope; the outer object
adds source styles/worksheet SHA-256 identities, warnings, nullable `text_rule`,
`data_bar` and `frozen_view` records. It does not create an editable workbook,
relax V1, repair the ZIP or evaluate formulas. V2 requires at least one qualified
conditional rule. The browser uses a separate V2 worker and closed decoder.

The closed initial profile admits at most one rule of each kind:

- `cellIs/equal` with an uppercase ASCII word literal (1–64 letters), DXF 0,
  explicit bold/RGB font and a background-only patternFill. Each input must be a
  stored uppercase ASCII string; normal formula caches are marked as cached.
  Each effect retains its source cell, saved value and exact-match decision.
  Background-only differential fill selection is a disclosed compatibility
  approximation; it does not alter the strict base-style resolver.
- A data bar over literal nonnegative integers inside explicit integer bounds
  (0–1,000,000,000), with explicit 0–100 min/max lengths and visible values.
  Legacy and x14 rules must agree on bounds, lengths, range and GUID ownership.
  Only the qualified gradient/axis-none form is admitted; its inactive negative
  color must match the bar and its inactive axis color must be black. Formula
  inputs, dynamic thresholds and competing extension declarations refuse.
  Each effect retains its literal value and proportional length percentage.

Rule priorities must be unique. Ranges must be disjoint, fully populated, within
V1's 128-row/32-column/4096-cell bounds and outside every merge. Data-bar cells
must have no base fill. Unknown metadata, unsupported DXFs, malformed links,
missing caches, noncanonical numeric inputs and partially qualified rules
refuse the whole result. No partial overlay is returned.

The proportional data-bar length follows Microsoft's [data-bar contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.databar?view=openxml-3.0.1)
conceptually: minLength + (value-min)/(max-min) * (maxLength-minLength). The
[Office 2010 contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.databar?view=openxml-3.0.1)
defines linked legacy/x14 declarations. The percentage is not a claim about
Excel pixel quantization, font metrics, printer scaling or pagination.

One frozen-row viewport form is separately qualified: zero frozen columns,
1–127 frozen rows, an A-column origin immediately below them, bottomLeft active
pane, and the two explicitly recorded A1 selections. `frozen_view` retains row
count, origin and active pane; a warning states that this viewport state is not
applied to the source grid. Other pane forms remain unsupported.

Tests use a synthetic merged-header workbook with two conditional cells,
exercise linked-rule/DXF/range/cache/viewport refusals, and assert V1, strict
extraction and mutation still refuse the conflicting source. Optional
`XLSX_SOURCE_CONDITIONAL_EVIDENCE_DIR` exports only that synthetic fixture and
JSON for worker/browser contract tests. External workbooks stay local.

The playground first tries V1, then V2 if V1 refuses. Each attempt owns and
terminates its worker; generation checks guard retries, cancellation, source
replacement and unmount. Only V2 effects joined to the decoded source cells may
change the preview: matched status cells receive the declared font/background
colors and bold text, while data bars sit behind their unchanged saved values.
Warnings retain the approximate geometry, cached-result and viewport omissions.
No native editing or download authority is exposed by either preview client.

The synthetic Chrome qualification runs both profiles. The conditional case
checks actual V1 refusal/V2 fallback, source hash, merge/header colors, saved
status/number values, differential colors and bar width. Cancellation and source
replacement wait for an actual V2 inspection before delivering delayed replies.
Screenshots are retained in the `xlsx-source-conditional-browser` CI artifact.
This does not replace local unchanged-source browser review or an independent
Excel print calibration matrix.
