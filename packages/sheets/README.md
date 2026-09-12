# @injoffice/sheets

Dependency-free contracts for fail-closed native spreadsheet saves. Version 1
normalizes editor changes into a small JSON vocabulary; it does not apply the
changes to XLSX bytes. `go/xlsxpatch` is the native application boundary.

```bash
npm install @injoffice/sheets
```

```ts
import { decodeWorkbookMutationBatch, encodeWorkbookMutationBatch } from '@injoffice/sheets'

const batch = {
  protocol: 'injoffice.xlsx.mutations',
  version: 1,
  batch_id: 'save-0001',
  expected_revision: 'opaque-revision-from-gateway',
  operations: [
    {
      operation_id: 'edit-0001',
      kind: 'cell.set_value',
      sheet_id: 'stable-sheet-id',
      cell: { row: 1, column: 2 },
      value: 42,
    },
  ],
} as const

const decoded = decodeWorkbookMutationBatch(batch)
if (!decoded.ok) {
  // Send issues directly in a 400 response or render them in the editor.
  console.error(decoded.issues)
} else {
  const body = encodeWorkbookMutationBatch(decoded.value)
  await fetch('/v1/xlsx/mutations', { method: 'POST', body })
}
```

`POST /v1/xlsx/mutations` is served by optional `injoffice-server` and by the
local `go run ./go/xlsxpatch/cmd/xlsxnative serve` helper. Both wrap
`go/xlsxpatch` extract/apply; the helper does not store artifacts.
`@injoffice/xlsx-wasm` packages the same functions for browser-local
extract/apply in a Worker. The playground uses that local path by default and
offers HTTP only as an explicit fallback.

Browser-local contract and mutation clients can import
`@injoffice/sheets/browser`. That narrow entry omits the root package's
Node-qualified HarfBuzz paint implementation.

## Version 1 wire semantics

- All coordinates are zero-based. Range ends are inclusive.
- `sheet_id` is the stable identity imported from the workbook. A sheet name
  or positional index is not a substitute.
- `batch_id` is a required idempotency key for one logical save. It is stable
  across retries and distinct from `expected_revision`. A gateway may return
  the recorded result of a previously successful batch with the same id
  instead of applying it twice.
- `expected_revision` is an opaque gateway-issued revision of the original
  XLSX bytes. The future application endpoint must compare it atomically and
  reject a stale batch before changing any bytes.
- `operation_id` is stable across retries, unique within one batch, and uses
  the restricted ASCII identifier vocabulary enforced by validation.
- Operations are applied atomically in array order; a later operation may
  intentionally overwrite an earlier operation in the same batch.
- `cell.set_value` writes a literal and removes any formula.
  `cell.set_formula` writes an A1 formula including `=`, removes any literal,
  and invalidates the prior cached result. `cell.clear_value` clears a literal;
  `cell.clear_formula` clears both the formula and cached result. Both clear
  operations leave the cell blank. Dates are Excel serial numbers plus a
  `number_format` style delta.
- A style property omitted from `style.patch` is unchanged. `null` clears
  that direct property. Unknown style keys are rejected.
- JSON numbers must be finite. Coordinates and dimensions are bounded by
  Excel limits. Unknown fields and operation kinds are rejected.

Version 1 deliberately refuses sheet add/delete/rename/reorder, row and
column insert/delete/move, and range moves. These kinds are exported in
`UNSUPPORTED_STRUCTURAL_MUTATION_KINDS`; validation returns
`UNSUPPORTED_OPERATION` instead of silently dropping them. Merges and
unmerges are the only supported structural changes.

Validation failures are JSON-serializable `WorkbookMutationIssue` values with
a stable code, RFC 6901 path, message, and operation index/id when available.

## Native XLSX import

`decodeNativeWorkbookV1` is the untrusted JSON boundary for the native
`go/xlsxpatch` extraction contract. Its TypeScript wire types and runtime shape
schema are generated from `schemas/xlsx-native-v1.schema.json`; a generated Go
binding manifest and reflection test keep field names, requiredness, and
resource limits synchronized with the canonical Go structs.

`decodeNativeWorkbookV2` is a separate protocol (`version: 2`, media type
`application/vnd.injoffice.xlsx-native.v2+json`) generated from
`schemas/xlsx-native-v2.schema.json`. v1 remains lockstep. The v2 extract
capability inventory is `native-ooxml-parse`/`read-only`, `native-geometry`/`exact`,
`native-decorations`/`exact`, `native-grid-commands`/`exact`, and
`unsupported-content`/`preserve-exact`. Extract still does not advertise
workbook-wide cell glyph/display paint, number-format display, or host fonts.
`compileNativeSheetCellPaintV2` is a separate producer: it paints producer-issued
display strings the exact Normal-font bytes can shape, with
`native-cell-glyphs`/`exact` on the paint plan, and records every refused cell
in that plan's `unsupported` list. The original XLSX remains the readable
authority for unpainted text. `projectNativeWorkbookV2` brands
`injoffice.xlsx.render-model` version 2.

`projectNativeWorkbookV1` produces a sparse, deeply frozen and privately branded
`NativeWorkbookRenderModelV1` for renderers and document pipelines. Geometry
and decoration compilers reject mutable or handcrafted lookalikes. The model
deliberately has no DOM, HTML,
Univer, React, Konva, Electron, or browser API dependency. Numeric and date
values remain OOXML lexical strings. Formula text and cached values remain
separate and are never evaluated. Absent `t` and explicit `t="n"` remain
distinct.

The projection retains stable document/sheet IDs, exact source revision and
package fingerprints, style IDs plus effective-style provenance, sheet/cell
editability, capability inventory, passthrough part fingerprints, and every
unsupported-content authority record. Unsupported subtrees in modeled core
workbook/worksheet/style/string parts remain authoritative in the source
package even though those parts are not listed as passthrough parts. Consumers
must honor `editable`, `refusal_code`, and `unsupported`; they must not infer
editability from rendered appearance.

The decoder rejects duplicate JSON keys, unknown fields, invalid unions or
lexicals, inconsistent style/formula references, forged refusal authority,
noncanonical OPC identities, and contracts exceeding the canonical resource
budgets. A validation failure never returns a partial render model.

JSON Schema `minLength` and `maxLength` count Unicode scalar values. Fields
whose native Office limits are defined in UTF-16 code units additionally carry
and enforce `x-maxUtf16Length`; this covers sheet names, formula/string text,
number formats, and font names.

## Deterministic native sheet geometry

`compileNativeSheetGeometryV1` turns a bounded viewport into immutable integer-EMU
row/column bands and merged rectangles. `emitNativeSheetGeometryCommandsV1`
records the same geometry as renderer-neutral commands that a Canvas or other
host adapter can replay without making that host the workbook authority.

The compiler uses exact `sheetFormatPr` defaults plus explicit row/column
overrides. Excel column widths are font-relative, so callers inject a
maximum-digit-width record bound to the exact workbook revision, source package,
the built-in Normal style XF/font record, resolved font digest, 96-DPI measurement
basis, and deterministic provider revision. Commands use viewport-local
coordinates and identify the origin cell explicitly. Missing defaults or metrics,
source dimension extras, clipped merged ranges, `zeroHeight` without explicit row
visibility provenance, unsafe integer geometry, and oversized viewports are
refused. The package never substitutes CSS, DOM, Canvas text measurement, a
screenshot, or an Office/LibreOffice render.

## Exact fill and border decorations

`compileNativeSheetDecorationsV1` consumes that source-bound geometry and emits
immutable fill rectangles plus orthogonal border segments. Fill colors and
border colors must be direct opaque OOXML RGB. Each effective fill and border is
bound to its exact styles-table record ID and raw-record SHA-256, and the whole
effective projection carries a consumer-recomputed SHA-256 so a valid-looking
color or token edit cannot retain stale provenance. Border styles remain their
exact OOXML tokens (`thin`, `double`, `dashed`, and the other schema tokens); the
contract deliberately supplies no stroke width because OOXML does not define an
authoritative physical width for those names. Every border source carries its
cell reference, effective style ID, `<border>` record ID and raw-record SHA-256.
Coordinates remain integer EMU from the geometry contract end to end.

An identical shared edge is emitted once with both source records. Different
styles or colors meeting on the same edge are refused instead of applying an
undocumented precedence heuristic. A partial effective style remains eligible
when its fill and border lanes are complete; unrelated unsupported font-theme or
number-format metadata does not disable those exact lanes. Compilation refuses
unsupported fill/border lanes, theme/indexed/automatic border colors,
diagonal/start/end/vertical/horizontal borders, merged cells, row or column
style inheritance, conditional
formatting, tables, drawings, external links, formula groups, mismatched
document/revision/package/geometry authority, styled hidden or zero-size cells
whose edge semantics cannot be represented, and resource overflow. Emission and
replay validate the full protocol/version/source identity, bounded integer
coordinates, containment, resources, command grammar, and recomputed decoration
digest before their first host callback, so a refusal writes no partial command
stream.

This slice paints backgrounds and symbolic borders only. Merged-cell border
semantics, conditional/table styles, drawings, and exact host mapping of OOXML
border tokens remain consumer work or explicit future native contracts.

## Exact cell glyph/display paint

`compileNativeSheetCellPaintV2` consumes source-bound geometry plus the same
producer-issued font bytes used by `createNativeMaximumDigitWidthAuthorityV2`.
It shapes one resolved Normal-style face through the pinned
`@injoffice/font-metrics` HarfBuzz path, parses integer TrueType `glyf`
outlines from those bytes, and places glyph origins in integer EMU inside the
geometry cell rect. `emitNativeSheetCellPaintCommandsV2` /
`replayNativeSheetCellPaintCommandsV2` record renderer-neutral `fillGlyphPath`
commands; the host adapter only fills paths.

The qualified subset is fail-closed and viewport-bounded:

- producer-issued literal strings and numeric lexicals the Normal font can
  shape, including BMP Latin/Common Unicode present in that font's cmap
- boolean cells as `TRUE`/`FALSE` from the stored 0/1/true/false lexical
- error cells as the producer-issued error lexical (`#DIV/0!`, `#N/A`, …)
  when it is printable ASCII
- date/numeric display for exact OOXML `y`/`m`/`d`/`h`/`s` tokens plus `General`,
  fixed decimals (`0` through six decimal places), optional `#,##0` grouping,
  percentages (`0%`, `0.00%`), and explicitly quoted currency prefixes/suffixes
  (`"$"#,##0.00`, `0.00" €"`), applied without `Date`, `Intl`, host locale, or host timezone;
  decimal scaling/rounding uses bounded integer arithmetic, including percentages.
  Separators are deterministic comma grouping and decimal point, not a claim of
  host-locale display parity. Accounting padding, colors/conditions, multi-section
  formats, fractions, scientific notation, scaling commas and implicit currencies
  remain refused rather than displaying raw numbers as if formatting succeeded;
  `General` numbers keep their stored lexical; named months/days (`mmmm`,
  `dddd`) paint only from an explicit OOXML `[$-…]` calendar/locale that has a
  producer table; system `[$-F800]` / host-locale names stay refused
- formula cells only when a cached lexical is already modeled; formulas are
  never evaluated
- fonts whose resolved typeface matches exact TTF bytes (Normal face plus extra
  sha256-keyed font-byte authorities); a different family is never substituted
- direct `#RRGGBB` font color, `general`/`left`/`center`/`right` and
  `top`/`middle`/`bottom` alignment
- `wrap_text` using HarfBuzz advances of the already-qualified display
  string: greedy wrap at U+0020, otherwise a measured character break;
  the modeled display string is not rewritten. Lines that escape the existing
  geometry row height are refused rather than clipping or inventing row height
- wrap then integer-EMU shrink-to-fit `floor(lineBox/lineAdvance)` when both
  alignment flags are set; overflow after that scale is refused
- reconstructible rich runs from SST/`<r>` (text plus font name/bold/italic/size/RGB)
- merged text only from the top-left cell
- neighbor spill or host measurement is refused rather than clipped by a heuristic

Missing reconstructible rich runs, unnamed locale dates, theme fonts without
exact TTF bytes, missing or unflattenable composite glyphs, combining marks /
ZWJ / variation selectors that the font cannot shape, CSS pixels, and
DOM/canvas measurement are recorded as `unsupported` for that cell or refused
for the compile. The paint plan advertises `native-cell-glyphs`/`exact` only
for the cells it actually painted. Host fonts remain a non-goal. Go extract
does not advertise glyph-paint identity.

DOM/HTML, Canvas measurement, screenshots, PDF rendering, system fonts, and the
host locale are never workbook authority.

### Supplemental table preview

`nativeStoredRowPreviewV1` reads optional source-bound stored row geometry from
object inspection. The projection covers the first 32 rows of up to 64 sheets;
unavailable or uninspected sheets retain host preview sizes. It applies stored
fixed heights and explicit hidden-row flags only. Automatic heights, unknown
dimension metadata, default-hidden sheets and ambiguous rows remain unavailable.
Columns and text metrics are not inferred. The exact `x14ac:dyDescent` attribute
qualifies fixed-height semantics as specified by
[MS-XLSX](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f11dfda4-46de-4035-8418-d76b0d3898f1),
including its `customHeight` side effect; baseline positioning remains a visible
limitation. Valid `spans` hints do not change stored row geometry. No native
mutation or exact geometry compiler refusal is weakened.

After source-bound object inspection, `nativeTableFillPreview` returns a qualified
table background for a cell with default-fill provenance.
`nativeTableHeaderTextPreview` identifies cells eligible for white bold header
text, without replacing explicitly selected font styles. Both require the source
revision, worksheet part, zero-based coordinates and cell style ID. A fill record
ID alone is not default-format authority: the cell style must appear in the
engine's source-qualified eligible list. Neither changes the workbook or grants
edit support. Overlapping tables, missing provenance and unsupported
styles remain unpainted; table warnings describe the partial result.

`nativeTableNumberFormatPreview` projects explicit table/column differential
number formats onto source-qualified General-format cells. Explicit cell formats
(including an explicit General override) take precedence. Saved formula results
are displayed without recalculation. `formatNativeAccountingTextPreview` is a
separate, display-only fallback for bounded numeric sections and quoted literals;
it preserves decimal lexical precision and reports omitted accounting fill and
padding alignment. It does not widen the exact native glyph-paint contract.

`nativeTableBorderPreview` supports a measured subset of built-in Medium2:
1-point outer/horizontal borders using accent1 with +0.4 HLS tint, and a 3-point
double totals divider in accent1. Source-qualified default border styles are
required on both sides of an edge. Explicit, unknown, overlapping-table, named
style and unsupported differential border overrides remain unpainted. Custom
DXF fonts/fills/borders, unqualified totals font styling, accounting positioning and text
metrics are not fully reproduced. This is partial read-only presentation, not
an Office-fidelity claim or mutation authority.

`nativeTableTotalsTextPreview` identifies qualified Medium2 totals-row cells for
bold text. It reuses source-qualified default-font style IDs and requires explicit
totals evidence, an unambiguous table and the matching package revision. Explicit
cell fonts and totals DXF/named-style overrides remain unchanged. Font family,
color, sizing and metrics are not replaced by this helper.
