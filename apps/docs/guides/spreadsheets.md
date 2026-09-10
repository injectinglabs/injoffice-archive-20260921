# Spreadsheets / XLSX

Use `@injoffice/sheets/browser` for browser-safe native contracts and `@injoffice/xlsx-wasm` for source-bound extraction and writes. The same Go writer is available to server applications.

## Before you begin

Check [package availability](../getting-started/packages#check-a-release-before-adopting-it). If the XLSX Worker package is not published for your chosen release, use a built source checkout; do not substitute a different spreadsheet serializer.

Configure the three version-matched Worker assets using the [browser integration guide](../integration/browser). Your host supplies the selected file as `Uint8Array` and owns the download action.

## Change a cell against its exact source

<<< @/examples/xlsx.ts

This function edits A1 on the first extracted sheet and returns replacement bytes plus a freshly extracted model. It does not silently choose a different target if the source refuses that edit. In your acceptance logic, inspect the returned native sheet/cell for the requested value before presenting a verified download.

## Coordinates and revisions

- Spreadsheet row and column coordinates are zero-based. A1 is `{ row: 0, column: 0 }`.
- Use the extracted stable sheet ID, not its name or positional index, in mutations.
- The outer package revision is `workbook.source.package_sha256`.
- The native model also carries an inner revision; `adaptWorkbookMutationBatchV1` translates and checks the two contracts.
- Keep operation IDs stable across a retry. The browser Worker itself is not a durable idempotency store.

## What this path can write

The native Worker adapter supports cell value/formula operations, supported style patches, row heights, and column widths. It requires the native family order: cells, then styles, then layout. It refuses merge/unmerge even though the broader Sheets mutation schema recognizes them. Structural row, column, and sheet changes are not implied.

Setting a cell value removes its formula. Setting a formula does not turn the native writer into a spreadsheet calculation engine. See [formulas](spreadsheet-tools#formulas-and-calculation).

## Connect an editor or agent

For an editor, translate only supported edits to the public mutation batch and bind it to the source the user actually reviewed. Keep rendering state separate from file authority.

For an agent, use `@injoffice/agent-office/xlsx` with host `snapshot`, `preview`, and `apply` callbacks. Preview must operate on an isolated copy; apply must enforce source freshness atomically. See [agent integration](../agents/integration).

## References

- [XLSX Worker contract and limits](../reference/generated/packages/xlsx-wasm)
- [Sheets mutations, projections, geometry, and paint](../reference/generated/packages/sheets)
- [Native Go XLSX engine](../reference/generated/go/xlsxpatch)
