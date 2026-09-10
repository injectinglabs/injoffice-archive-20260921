import { adaptWorkbookMutationBatchV1, createXlsxWasmClient } from '@injoffice/xlsx-wasm'
import type { WorkbookMutationBatchV1 } from '@injoffice/sheets/browser'

/** Browser-only execution: the host owns file selection and download. */
export async function updateFirstCell(original: Uint8Array, value: string) {
  const client = createXlsxWasmClient()
  try {
    const workbook = await client.extract(original)
    const firstSheet = workbook.sheets[0]
    if (!firstSheet) throw new Error('The workbook has no sheets')
    const batch = {
      protocol: 'injoffice.xlsx.mutations', version: 1,
      batch_id: 'edit-first-cell',
      expected_revision: workbook.source.package_sha256,
      operations: [{
        operation_id: 'set-a1', kind: 'cell.set_value', sheet_id: firstSheet.id,
        cell: { row: 0, column: 0 }, value,
      }],
    } satisfies WorkbookMutationBatchV1
    const transaction = adaptWorkbookMutationBatchV1(workbook, batch)
    const output = await client.apply(original, workbook, transaction)
    const readback = await client.extract(output)
    return { output, readback }
  } finally {
    client.terminate()
  }
}
