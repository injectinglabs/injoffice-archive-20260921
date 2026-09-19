# Bounded XLSX conditional fill preview

The object-inspection API can project a read-only conditional fill supplement, and the spreadsheet page demo applies it only when **Preview supported conditional fills** is selected. For example, one `cellIs` rule `greaterThan 0` with an explicit green solid-fill differential style colors saved values `3` and `7`, while retaining the ordinary fill for `-2` and `0`.

This is a narrow source-based display feature. It does not change source styles, resolve general conditional formatting, recalculate formulas, authorize mutations or establish Excel print fidelity. The existing native conditional-formatting preservation inventory and table-style refusal policies remain unchanged.

## Qualified source profile

- Exactly one `conditionalFormatting` element and one `cfRule` in the worksheet. Misplaced/foreign competing rules, worksheet extension lists and alternate content refuse the overlay, even outside the selected preview range.
- One canonical uppercase A1 cell or rectangle in `sqref`, at most 4,096 cells. A workbook supplement holds at most 64 worksheet entries and 16,384 evaluated cells in total.
- Rule type `cellIs`; operator `equal`, `notEqual`, `lessThan`, `lessThanOrEqual`, `greaterThan`, or `greaterThanOrEqual`; exactly one literal integer operand. Lexicals use an optional minus followed by `0` or 1–15 digits without leading zeros. Negative zero compares as zero while retaining its source lexical. Decimal, exponent, leading-plus and leading-zero forms are outside the profile.
- An explicit positive priority and an optional schema boolean `stopIfTrue` (default false) are retained. With one rule, no priority conflict or stop behavior is inferred. Multiple rules always refuse, irrespective of ranges, priorities or whether their conditions would be true.
- An explicit `dxfId` into the workbook-related styles part. The selected differential style must contain only `fill/patternFill`, `patternType="solid"`, and one `fgColor` with opaque `FFRRGGBB`. Themes, indexed colors, alpha, tint, background colors, extra styling and extension content are outside this profile.
- Every targeted cell must have a saved numeric integer value in the same bounded lexical subset. An ordinary formula can supply its saved numeric cache; its freshness remains unknown. Any formula group on the worksheet refuses the overlay, including unmarked array followers. Missing cells, blanks, text, booleans, errors, absent caches, metadata and opaque cell markup are not coerced.
- Target merges and worksheets containing table parts refuse the overlay. Existing table-style preview logic stays unchanged.

The producer validates the complete rule and target range before emitting any matches. An unsupported case yields a worksheet `unavailable` entry and a specific warning, without a partial rule or cell overlay. Base page previews remain usable with that omission visible.

## Public API and source ownership

`InspectNativeWorkbookObjectsV1` adds optional `conditional_fills` to the existing `injoffice.xlsx.preview-objects` version 1 supplement. The Go/WASM inspect path exposes the same field. Old supplements without it remain accepted. The mutation contracts and generated native workbook schemas are unchanged.

`decodeNativeConditionalFillPreviewsV1` validates closed entry shapes, source sheet identities, bounds, complete row-major cell coverage, integer comparisons, styles and warnings. `decodeNativeWorkbookObjectsV1` validates the additive field within its package hash envelope. `selectNativeConditionalFillPreviewV1(workbook, sheetId, objects)` additionally joins every stored lexical/cache owner and the worksheet identity to the opened workbook, refusing stale values, groups, merges, duplicate identities or unsupported cell inventory. Consumers should use the selector before drawing an overlay.

```ts
import { selectNativeConditionalFillPreviewV1 } from '@injoffice/sheets/browser'

const preview = selectNativeConditionalFillPreviewV1(workbook, sheetId, objects)
if (preview?.status === 'available') {
  const matches = preview.cells.filter(cell => cell.matches)
  // Apply preview.rule.fill only to these source-joined cells.
  // Show preview.warnings, including saved-cache limitations.
}
```

The page demo exposes the source range, operator, operand, priority, `stopIfTrue`, differential style index, fill and matching/cached counts. Source changes reset the opt-in choice, and changing it clears stale pages. The overlay overrides the ordinary cell fill only for matching cells; no workbook field is rewritten.

## Normative basis and qualification limits

Microsoft's [conditional-formatting documentation](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting) describes worksheet-owned `sqref` ranges, static `cellIs` operands, differential style references and worksheet-global priorities. The [rule definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.conditionalformattingrule) describes the comparison and `stopIfTrue` attributes. [Pattern fills](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternfill) define the solid foreground fill, and [Excel's precedence guidance](https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel) specifies that a true conditional rule takes precedence over the existing manual format.

Generated software tests and browser/WASM evidence demonstrate this implementation profile. They are not independent Excel output comparisons. General rule evaluation, overlapping priorities, richer differential styles, decimal/coercion behavior, table/merge combinations and measured print calibration remain outstanding. Source files, fonts, screenshots and qualification scripts belong in the local evidence corpus, not this repository.

## Colour-scale fills

A second, independent tier projects `colorScale` rules. Unlike the `cellIs` overlay above it is not opt-in: a colour scale is the cell's only visible content in Excel, so it is applied wherever the source qualifies, in the same place a table-style fill is applied. It paints *behind* the cell, so an explicit source fill or a table-style fill always wins.

### Qualified source profile

- Two or three `cfvo` bounds with a matching number of `color` stops, in that order and with nothing else inside `colorScale`.
- Bound types `min`, `max`, `percent`, `percentile` and `num`. A `num` bound is a decimal literal or a reference to one cell on the same worksheet (`$A$1`); anything else, including `formula`, leaves that rule unpainted. Bounds must be ascending once resolved.
- Stop colours are an opaque `FFRRGGBB` or a theme slot with an optional tint, resolved through the same theme table the styles tier uses. Indexed and translucent colours are outside the profile.
- Every painted cell must carry a saved numeric value. Blanks, text, booleans and errors take no fill and are excluded from the scale's own minimum, maximum and percentiles, which is what Excel does.
- A range that meets any other rule on the worksheet — including a rule carried in an `x14` worksheet extension, and including one this tier cannot read — is left unpainted, because the result then depends on priority and `stopIfTrue`, which are not evaluated here. A range that meets a merged region is left unpainted.
- At most 4,096 cells per range, 16,384 per workbook and 64 worksheet entries.

A rule outside the profile costs only itself: the other colour scales on the sheet are still painted and the count of unpainted rules is disclosed in the entry's warning. A sheet where nothing resolved reports `unavailable` with a reason.

### Public API

`InspectNativeWorkbookObjectsV1` adds optional `conditional_scale_fills` to the `injoffice.xlsx.preview-objects` version 1 supplement; supplements without it remain accepted. `decodeNativeConditionalScaleFillPreviewsV1` validates closed entry shapes, sheet identities, colours and row-major cell ordering. `nativeConditionalScaleFillPreview(objects, revision, sheetPart, row, column)` returns the colour for one zero-based coordinate, and returns nothing when the supplement does not join the opened package hash.

Colour interpolation is channel-wise in sRGB between the two bounds that surround the value, clamped outside the end bounds; a three-stop scale interpolates each half separately. Percentile bounds use `PERCENTILE.INC` over the range's own numeric values.

This tier changes no source style, recalculates no formula, authorizes no mutation and does not remove the worksheet's `CONDITIONAL_FORMATTING` preservation entry. `dataBar` and `iconSet` rules remain unpainted.

## Data bars

A third tier projects `dataBar` rules, applied the same way a colour scale is: over the cell's own fill and under its text, which is the order Excel paints them.

### Qualified source profile

- The downlevel `dataBar` element supplies the positive fill colour; everything that decides the geometry is read from the `x14` twin the rule names through its `{B025F937-C7B1-47D3-B67F-A62EFF666E3E}` extension. A rule with no twin is not painted, because its axis, negative colour and gradient flag are then unknown.
- `gradient="0"`, `minLength="0"` and `maxLength="100"` — the flat, unclipped bar. A gradient bar's colour varies along its own length, and a clipped length moves every bar's start and end, so both are left to a later tier rather than approximated.
- `axisPosition` absent, `automatic` or `none`. A `middle` axis is not painted.
- Bound types `autoMin`, `autoMax`, `min`, `max`, `num`, `percent` and `percentile`, resolved exactly as a colour scale resolves them; `autoMin` and `autoMax` are Excel's automatic bounds, `min(0, lowest)` and `max(0, highest)`. A `formula` bound leaves the rule unpainted.
- Colours are an opaque `FFRRGGBB` or a theme slot with an optional tint, for the bar, the negative fill, the border, the negative border and the axis. A rule that declares `border="1"` without a border colour is not painted.
- The same neighbour, merge and budget rules as the colour-scale tier apply. A rule owns its range twice here — once downlevel and once in the extension — so a range is only painted when exactly those two claims cover it.

### Geometry

The whole bound span maps onto the cell's width. The axis is wherever zero falls inside that span, each bar runs from the axis to its own value, and a value below the axis is drawn to the left of it in the negative colour. The projection reports the span and the axis as thousandths of the cell's width, so the renderer needs no source units. Excel's own few-pixel inset inside the cell is not reproduced: the bar spans the cell.

### Public API

`InspectNativeWorkbookObjectsV1` adds optional `conditional_bar_fills` to the `injoffice.xlsx.preview-objects` version 1 supplement. `decodeNativeConditionalBarFillPreviewsV1` validates closed entry shapes, ordered coordinates, in-cell spans and colours, and refuses an axis position without its colour or a colour without its position. `nativeConditionalBarFillPreview(objects, revision, sheetPart, row, column)` returns the bar for one zero-based coordinate.

`iconSet` rules remain unpainted.
