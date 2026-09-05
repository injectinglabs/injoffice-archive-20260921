# xlsxpatch

`xlsxpatch` performs focused XLSX changes while keeping the original ZIP archive as the source of truth. Untouched parts are copied and verified instead of reconstructing the workbook through a partial object model.

```bash
go get github.com/injectinglabs/injoffice/go/xlsxpatch
```

```go
package main

import (
	"fmt"
	"os"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func main() {
	data, err := os.ReadFile("input.xlsx")
	if err != nil { panic(err) }
	inspection, err := xlsxpatch.Inspect(data)
	if err != nil { panic(err) }
	fmt.Printf("%+v\n", inspection)
}
```

The module can read and write charts and shapes, hydrate/add/update/remove native pivot parts, hydrate and apply print setup, and apply explicit archive part patches. Chart lifecycle calls use chart-part plus drawing-part/`cNvPr` object id as stable identity; update/delete preserve sibling drawing objects and refuse shared parts, unsupported dependency graphs, ambiguous ownership, and stale identities. Shape lifecycle calls use drawing-part plus `cNvPr` object id as stable identity; update/delete replace or remove one top-level anchor, preserve sibling objects and connector bindings, and reject duplicate ids, grouped targets, ambiguous ownership, stale identities, and dangling connector deletes. Group editing remains outside the bounded API. Pivot lifecycle calls use the hydrated pivot-table part as stable identity, preserve shared caches, and reject ambiguous relationship graphs. `ReadPrintSetups` reads the bounded worksheet page setup, print options, margins, centered/raw odd header/footer text, print area, and repeated-title defined names while reporting unsupported native settings. `SetPrintSetup` surgically writes print areas, named paper sizes, portrait/landscape orientation, fit or explicit scale, margins, centering, gridlines, headings, and centered header/footer text. Repeated-title write-back, custom paper, watermarks, and complete header/footer syntax remain outside that bounded writer. Always handle returned errors and hydration warnings; they are part of the fail-closed contract.

## Native worksheet outlines

`ReadWorksheetOutline` projects a sheet's row/column `outlineLevel`, `hidden`,
and `collapsed` metadata plus `sheetPr/outlinePr` summary direction. It derives
properly nested/disjoint groups with deterministic IDs and rejects malformed,
ambiguous, crossing-equivalent, or orphan collapsed state. Omitted
`summaryBelow` and `summaryRight` attributes use the SpreadsheetML default of
`true`.

`ApplyWorksheetOutline` atomically replaces outline-owned metadata for one
stable `sheetId`, then reopens the result and verifies the reconstructed group
projection. Unrelated OPC parts are raw-copied through `Apply`; opaque worksheet
content and independently hidden level-zero rows/columns remain intact.
SpreadsheetML can store seven nested group levels (`outlineLevel` 1 through 7),
so an eighth nested renderer-level group is explicitly refused at this native
boundary.

## Native XLSX extraction

`ExtractNativeWorkbookV1` is the additive, renderer-free import boundary for
workbooks. It reads SpreadsheetML and OPC directly; it does not render HTML,
invoke LuckyExcel, evaluate formulas, or regenerate the archive.

`ExtractNativeWorkbookV2` is a separate protocol (`version: 2`, media type
`application/vnd.injoffice.xlsx-native.v2+json`). It reuses the lockstep v1
extractor and emits the v2 schema digest plus this ordered capability
inventory:

- `native-ooxml-parse` / `read-only`
- `native-geometry` / `exact` — integer-EMU row/column bands and merged rects
- `native-decorations` / `exact` — fills plus producer-issued orthogonal stroke/border segments
- `native-grid-commands` / `exact` — renderer-neutral geometry and decoration command replay
- `unsupported-content` / `preserve-exact`

v1 remains the mutation/save contract. v2 does not claim cell glyph/display
paint, number-format display, or host fonts; the original XLSX remains the
readable authority for text.

```go
nativeV2, err := xlsxpatch.ExtractNativeWorkbookV2(input)
if err != nil { panic(err) }
```

```go
native, err := xlsxpatch.ExtractNativeWorkbookV1(input)
if err != nil { panic(err) }

// Keep document identity stable while importing a later exact-byte revision.
next, err := xlsxpatch.ExtractNativeWorkbookV1WithOptions(updatedInput,
	xlsxpatch.NativeWorkbookExtractionOptions{Previous: native})
```

The v1 contract uses the workbook's unique `<sheet sheetId>` as the stable
`sheet_id` accepted by the native mutation appliers. It exposes sheet
name/order/state, sparse cells, row heights, canonical column ranges, raw cell
style IDs, canonical read-only merged ranges, exact `sheetFormatPr` default
geometry when present, the built-in Normal style's source font identity when
unambiguous, and an effective projection of the supported `style.patch` fields.
Effective styles resolve the OOXML `applyFill`/`applyBorder` joins between
`cellXfs` and `cellStyleXfs`. Each style carries a canonical effective-projection
SHA-256. A supported solid direct-RGB fill is bound to its exact `<fills>` index
and raw-record SHA-256. A supported border is bound to its exact `<borders>`
index and raw-record SHA-256 and projects only direct opaque RGB left/right/top/bottom
sides with a standard OOXML style token. Empty borders are exact. Theme,
indexed, automatic, diagonal, start/end, vertical/horizontal, foreign, or
otherwise ambiguous border records are marked `STYLE_BORDER`, make the style
projection partial, and remain source authority. Border mutation is not part of
v1; existing border records and references are preserved across supported
font/fill/alignment/number-format mutations.

An explicit `<numFmt>` declaration is source authority even when its ID is in
the nominal built-in range. It overrides the built-in fallback for that ID; the
overridden fallback code is not left aliased to the authored record and receives
a non-colliding custom ID if a later mutation requests it.

A merged range retains its exact inclusive
grid coordinates and canonical A1 reference; its top-left cell is the content
authority while blank covered cells and their independent style IDs remain
explicitly distinguishable. Hidden rows and columns do not renumber the range.
Numeric/date/error values remain lexical strings. Shared and inline strings are
decoded with SpreadsheetML `ST_Xstring`; formula text and typed cached values
are returned without calculation or coercion. Shared, array, and data-table
formula groups plus rich strings are readable but explicitly mutation-refused.

`revision`, package fingerprints, and passthrough fingerprints contain full
SHA-256 values. `capabilities`, `passthrough_parts`, and `unsupported` use the
same preservation-policy shape as the DOCX native contract. Unmodeled charts,
pivots, drawings, media, external links, extensions, and opaque parts remain
authoritative in the original XLSX and are inventoried with their actual ZIP
spelling, effective content type, byte length, and exact hash.

`source.authority="exact-package-bytes"` applies to both passthrough parts and
unsupported subtrees or attributes inside modeled workbook, worksheet, styles,
and shared-string parts. `passthrough_parts` is therefore not the complete list
of source-authoritative content: callers must also enforce the scoped
`unsupported` inventory against the matching `source.package_sha256` and
`revision` before permitting a mutation.

Extraction fails closed on ambiguous or nonconforming OPC routing, case- or
percent-equivalent duplicate parts, wrong/missing content types, non-exact
`TargetMode`, Strict/Transitional disagreement, unsafe style tables, invalid
cell/shared-string/style references, malformed, duplicate, overlapping, or
out-of-bounds merged ranges, covered non-anchor content, overlapping column
ranges, hostile XML, and bounded archive/XML/cardinality/resource limits.
Merged-range mutation remains refused until an exact atomic native writer is
proven. `Previous` is validated as a complete v1 contract for the same canonical
workbook and cannot remap an existing sheet identity to another worksheet part.

## Atomic native mutation transaction

`ApplyNativeWorkbookMutationTransactionV1` is the save boundary that composes
the cell/formula, bounded style, and row/column appliers without reconstructing
the workbook:

```go
result, err := xlsxpatch.ApplyNativeWorkbookMutationTransactionV1(input,
	xlsxpatch.NativeWorkbookMutationTransactionV1{
		ExpectedRevision: native.Revision,
		Cells: []xlsxpatch.CellMutation{{
			OperationID: "edit-total", SheetID: "7",
			Kind: xlsxpatch.CellSetFormula,
			Cell: xlsxpatch.CellRef{Row: 4, Column: 2},
			Formula: "=SUM(C1:C4)",
		}},
	})
if err != nil { panic(err) }
savedBytes, reopened := result.Package, result.Workbook
```

The expected revision must be the full SHA-256 revision emitted for the exact
input package; stale saves are rejected. Before applying anything, the whole
transaction is bounded and validated, operation IDs must be unique across all
three families, and every target is checked against the extractor's strict
sheet/cell refusal, merged-range authority, and style-projection state. The
merged-range check covers sparse/blank cells with a bounded row sweep rather
than expanding merged rectangles. Style ranges are additionally limited to
10,000 cumulative cell visits before expansion. An unsupported or oversized
later operation therefore returns no result and cannot commit earlier
operations.

Envelope handlers should use `ApplyNativeWorkbookMutationPayloadV1`, which
accepts only a bounded single JSON object. Its decoder rejects duplicate keys
at every nesting level, unknown fields, excessive depth/token counts, and
trailing JSON. It also requires an exact, unambiguous CAS join: the outer
`sha256:<digest>` revision must fingerprint the input bytes and the native
payload must carry the corresponding `rev:<digest>` value.

`POST /v1/xlsx/extract` and `POST /v1/xlsx/mutations` are served by the
optional `injoffice-server` and by the loopback helper. Both use the shared
`xlsxhttp` handlers; the helper does not store artifacts:

```bash
go run ./go/xlsxpatch/cmd/xlsxnative serve --addr 127.0.0.1:18765
```

```bash
go run ./go/xlsxpatch/cmd/xlsxnative extract input.xlsx > native.json
go run ./go/xlsxpatch/cmd/xlsxnative apply original.xlsx payload.json > saved.xlsx
```

`serve` binds localhost only and has no auth. `POST /v1/xlsx/extract` accepts a
raw XLSX body or multipart file and returns native v2 JSON.
`POST /v1/xlsx/mutations` accepts multipart `original`, `payload`, and
`expected_revision` (`sha256:<digest>`) and returns saved XLSX bytes.

The optional `@injoffice/xlsx-wasm` npm package wraps the same extract/apply
functions for local browser workflows. It is not the default playground path
and does not provide storage or collaboration:

```bash
./scripts/build-xlsxnative-wasm.sh
```

JS bindings: `xlsxnative.extract(bytes) -> json`,
`xlsxnative.apply(original, payload, expectedRevision) -> bytes`. See
`packages/xlsx-wasm/README.md` and `cmd/xlsxnativewasm/README.md`.

On success, the package is reopened through the native extractor using the
previous contract, contract validation is rerun, requested value/formula,
style, and layout postconditions are checked, and lexical content of cells not
targeted by value/formula edits is compared. Existing style records, cell style
assignments outside style targets, untouched cell edit authority, the complete
merged-range projection, and row/column projections outside their layout
targets must also remain exact. Every unrelated OPC entry is verified to retain
its exact compressed bytes and ZIP metadata. `Package` remains the save
authority; `Workbook` is the validated readback, not a model from which callers
may regenerate XLSX bytes.

Successful saves must make an authoritative semantic byte change; already
satisfied transactions are refused. Sheet topology and mutation-refusal state,
the complete unsupported inventory, and the cell set are checked in both
directions. Style clears are additionally verified against the reopened OOXML
style table so inherited values and cleared direct component flags, not merely
the flattened effective projection, are postconditions.

## Native cell values and formulas

`ApplyCellMutations` implements the value/formula subset of
`@injoffice/sheets` protocol v1 without rebuilding the workbook:

```go
output, err := xlsxpatch.ApplyCellMutations(input, []xlsxpatch.CellMutation{
	{
		OperationID: "edit-1",
		SheetID:     "7", // <sheet sheetId="7">; stable across rename/reorder
		Kind:        xlsxpatch.CellSetFormula,
		Cell:        xlsxpatch.CellRef{Row: 4, Column: 2},
		Formula:     "=SUM(A1:A4)",
	},
})
```

The v1 cell applier accepts strings, booleans, finite numbers, and A1 formulas.
It removes stale cached values, preserves existing cell styles and `extLst`
markup, applies SpreadsheetML `ST_Xstring` escaping, expands an existing
worksheet dimension, and removes a stale row `spans` hint when inserting cells.
Every accepted mutation marks the workbook for full recalculation. An existing
calculation chain and its relationship/content-type registration are removed
atomically because literal changes can invalidate dependent formula caches.

Untouched OPC parts are raw-copied and verified byte-for-byte. Array/shared/data
table formula ranges, merged-cell interiors (Transitional and Strict OOXML),
modern `cm`/`vm` cell metadata, unsupported cell children, malformed
coordinates/XML, unknown mutation kinds, or ambiguous package entries refuse
the entire batch and return no output.

## Native row heights and column widths

`ApplyLayoutMutations` implements exactly `row.set_height` and
`column.set_width` from `@injoffice/sheets` protocol v1:

```go
output, err := xlsxpatch.ApplyLayoutMutations(input, []xlsxpatch.LayoutMutation{
	{
		OperationID:  "resize-row",
		SheetID:      "7",
		Kind:         xlsxpatch.RowSetHeight,
		Row:          4, // zero-based
		HeightPoints: 22.5,
	},
	{
		OperationID: "resize-column",
		SheetID:     "7",
		Kind:        xlsxpatch.ColumnSetWidth,
		Column:      2, // zero-based
		Width:       12.25,
	},
})
```

Operations are validated as one atomic ordered batch, so repeated writes to the
same row or column use the last value. A zero dimension hides its row or column;
a positive dimension explicitly unhides it. Existing rows are changed only in
their start tags, missing rows are inserted in numeric order, and optional row
numbers are inferred sequentially. Existing canonical, non-overlapping `col`
ranges are split only where necessary while unrelated attributes are retained.
An absent `cols` element is inserted immediately before `sheetData`, as required
by SpreadsheetML schema order. Already-satisfied dimensions are byte-level
no-ops, including explicit `<col></col>` serialization.

## Native cell styles

`ApplyStyleMutations` implements exactly `style.patch` from
`@injoffice/sheets` protocol v1. Omitted properties remain unchanged; a null
property clears the direct value so the effective `cellStyleXf` inheritance is
used:

```go
output, err := xlsxpatch.ApplyStyleMutations(input, []xlsxpatch.StylePatchMutation{
	{
		OperationID: "format-total",
		SheetID:     "7",
		Kind:        xlsxpatch.StylePatch,
		Range:       xlsxpatch.StyleRange{Row: 4, Column: 1, EndRow: 4, EndColumn: 3},
		Style: xlsxpatch.StyleDelta{
			NumberFormat: xlsxpatch.SetStyleProperty("$#,##0.00"),
			Bold:         xlsxpatch.SetStyleProperty(true),
			FillColor:    xlsxpatch.SetStyleProperty("#DDEEFF"),
			WrapText:     xlsxpatch.ClearStyleProperty[bool](),
		},
	},
})
```

The bounded, ordered applier supports number format; font name, size, bold,
italic, and RGB color; solid RGB fill; horizontal and vertical alignment; and
wrapping. It creates missing rows/cells in canonical order and clones or
deduplicates fonts, fills, number formats, and cell formats without mutating a
shared style record. Formulas, values, opaque cell children, unrelated style
properties, and untouched ZIP parts are preserved. A semantic no-op returns a
clone of the original XLSX bytes without rewriting the archive.

Style-table inheritance follows the component-specific OOXML `apply*`
defaults. Unsafe references, ambiguous style tables, unsupported clear/inherit
through gradient, pattern, theme, or indexed fills, malformed XML, and an
unbounded range refuse the entire batch. An explicit RGB font or fill replaces
the targeted color component without flattening the shared source record.

The native mutation appliers route stable sheet IDs through the package's root
`officeDocument` relationship instead of assuming `xl/workbook.xml`. Relocated
workbooks and Transitional or Strict relationship types are supported. OPC part
lookups use decoded ASCII case-insensitive URI equivalence while preserving the
original ZIP entry spelling. Ambiguous equivalent package parts, external or
malformed routing relationships, overlapping/noncanonical column definitions,
unsafe XML, non-finite or out-of-bounds dimensions, and unsupported payloads
refuse the complete batch and return no output.
`ReadSparklines` and `SetSparklines` provide bounded, fail-closed persistence
for Excel's standard x14 worksheet sparkline extension. They support line,
column, and win/loss groups and the options represented by
`@injoffice/sparklines`; unsupported extension state is diagnosed and blocks
destructive replacement. Writes change only affected worksheet parts, retain
unrelated `extLst` children, and use the package's byte-preserving `Apply`
path for every untouched OPC entry.

`ReadConnectorDefinitions` and `SetConnectorDefinitions` persist InjOffice's
credential-free connector model in the explicitly versioned custom OPC part
documented in `docs/XLSX-CONNECTOR-EXTENSION.md`. Add/remove operations change
only the owned part, its workbook relationship, and its content-type override;
updates replace only the owned part. Unknown versions, fields, owners, targets,
or relationship graphs fail closed, while unrelated custom XML and package
parts remain byte-identical. This is an InjOffice round-trip extension, not a
claim of Excel Power Query or external-connection interoperability.
