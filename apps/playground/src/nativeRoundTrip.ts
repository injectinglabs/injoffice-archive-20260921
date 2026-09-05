import {
  WORKBOOK_MUTATION_PROTOCOL,
  WORKBOOK_MUTATION_VERSION,
  assertNativeWorkbookV2,
  type NativeWorkbookCellV2,
  type NativeWorkbookSheetV2,
  type NativeWorkbookV2,
  type WorkbookMutationBatchV1,
} from '@injoffice/sheets/browser'

export type NativeCell = NativeWorkbookCellV2
export type NativeSheet = NativeWorkbookSheetV2
export type NativeWorkbook = NativeWorkbookV2

export type EditableTarget = NativeCell & { sheetId: string; sheetName: string }

export function decodeNativeWorkbook(value: unknown): NativeWorkbook {
  assertNativeWorkbookV2(value)
  return value
}

export function editableTargets(workbook: NativeWorkbook): EditableTarget[] {
  return workbook.sheets.flatMap((sheet) => {
    if (!sheet.editable || !Array.isArray(sheet.cells)) return []
    return sheet.cells
      .filter((cell) => cell.editable && !cell.formula)
      .map((cell) => ({ ...cell, sheetId: sheet.id, sheetName: sheet.name }))
  })
}

export function targetKey(target: Pick<EditableTarget, 'sheetId' | 'row' | 'column'>): string {
  return `${target.sheetId}:${target.row}:${target.column}`
}

export function displayCellValue(cell: Pick<NativeCell, 'value' | 'formula'> | undefined): string {
  if (!cell) return ''
  if (cell.formula) return `=${cell.formula.text}`
  return cell.value?.text ?? cell.value?.lexical ?? ''
}

export function buildCellMutation(workbook: NativeWorkbook, target: EditableTarget, value: string, operationID: string): WorkbookMutationBatchV1 {
  return {
    protocol: WORKBOOK_MUTATION_PROTOCOL,
    version: WORKBOOK_MUTATION_VERSION,
    batch_id: operationID,
    expected_revision: workbook.source.package_sha256,
    operations: [{
      operation_id: operationID,
      sheet_id: target.sheetId,
      kind: 'cell.set_value',
      cell: { row: target.row, column: target.column },
      value,
    }],
  }
}

export function downloadName(sourceName: string): string {
  const base = sourceName.replace(/\.xlsx$/i, '').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '')
  return `${base || 'workbook'}-injoffice.xlsx`
}
