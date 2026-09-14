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
