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
