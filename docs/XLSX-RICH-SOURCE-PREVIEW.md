# Read-only rich source content

`xlsxpatch.PreviewNativeRichSourceV1(originalBytes)` returns a separate
`injoffice.xlsx.rich-source-preview` version 1 envelope. It never constructs an
editable native workbook, changes source XML/ZIP bytes, or grants mutation
capabilities. Native extraction, V1/V2 source-style recovery and mutation retain
their existing refusal behavior.

This profile handles an absent `cellStyleXfs` count by counting 1–128 original
direct records and reporting `parent_count: {table: "cellStyleXfs",
declaration: "absent", observed: ...}`. All other counts remain mandatory and
exact. Duplicate/foreign/malformed records, unknown style metadata, reference
errors and inheritance conflicts refuse the entire result. The envelope binds
the original package, workbook/styles/worksheet paths, styles/worksheet hashes,
used XF/parent/component hashes and strict registry error.

The source must contain one visible Transitional worksheet with a fully stored,
unmerged A1-based content rectangle, at most 128 rows, 32 columns and 4096 cells.
Every cell must use a nonzero style with explicit font/fill/border/number-format/
alignment ownership. Only used styles are projected; unused style IDs are
recorded. This permits an unresolved unused default font without inventing its
font or applying it to sparse cells. All style references still receive strict
validation. Used fonts need explicit qualified names, sizes and colors; direct
basic fills/thin RGB borders and simple alignment use the source-grid
qualifiers. Disabled strike/underline and left-to-right reading order are
separately qualified. Active effects, unresolved font color and formats other
than General refuse.

Values are inline strings or literal numbers. General numeric lexicals and
spaces stay exact; there is no floating-point display conversion, formula
execution or saved-cache freshness claim. At least one inline rich cell must
qualify completely through the existing direct-run parser. Every run joins the
extracted run text and concatenated cell text, with contiguous UTF16 start/end
offsets. Bounds: 64 runs and 2048 UTF16 units per rich cell, 256 rich cells, 1024
total runs, 65536 aggregate cell text/lexical units, 4096 bytes per cell text and
256 bytes per numeric lexical. Shared-string tables, theme run selection,
active effects, phonetic/unknown content and formula mixtures refuse.

Every displayed row/column needs explicit source geometry. Blank declared
bands outside the content rectangle are bounded and recorded separately as
`omitted_rows` (points) and `omitted_columns` (source width units), with zero-based
indices. Hidden bands, best fit, styled bands and missing displayed geometry
refuse. Empty outline properties and an empty `headerFooter` with
`alignWithMargins="0"` are qualified omissions. Viewport/print settings are not
applied. Browser font matching, shaping, wrapping and width metrics remain
approximate; the envelope is not an Excel print layout or a full worksheet view.

The browser package exports `createXlsxRichSourcePreviewClient` and
`decodeXlsxRichSourcePreviewV1`. The client snapshots source bytes and joins the
package SHA256 before returning a deeply frozen closed envelope. It exposes
only `preview` and `terminate`. Its self-contained `xlsxrichsource.worker.js`
loads a separate `xlsxrichsource.wasm` module with only `previewRichSource`; the
worker refuses extraction and mutation. The native module's size ceiling is
7.25 MiB + 8 KiB, read from `go/xlsxpatch/cmd/xlsxnativewasm/max-bytes.txt`; the
rich-source module retains a 7 MiB bound.
The native editing module and V1/V2 bindings remain compatible.

The playground tries V1, then V2, then rich-source recovery after native opening
fails. Every attempt is terminated, and generation guards cover cancellation,
replacement and unmount. Rich spans preserve joined source text and direct
properties with cell-font fallback; General numbers keep their exact lexicals.
Outside blank geometry and count/source evidence are disclosed in details.
Synthetic positive/refusal tests verify source bytes, strict refusals, exact
Unicode run joins, numeric lexicals and outside geometry. Actual WASM contract
checks cover rejected arguments, bound enforcement, recovery and Go parity;
synthetic browser checks cover the third-worker fallback and late replies.
Set `XLSX_RICH_SOURCE_EVIDENCE_DIR` when running `TestRichSourcePreview` to export
that synthetic workbook and its envelope for browser/decoder tests. Unchanged
external workbooks and independent reference PDFs remain local-only evidence.
