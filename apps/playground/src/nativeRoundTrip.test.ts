import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  buildCellMutation,
  decodeNativeWorkbook,
  displayCellValue,
  downloadName,
  editableTargets,
  targetKey,
  type NativeWorkbook,
} from './nativeRoundTrip'

const workbook = decodeNativeWorkbook(JSON.parse(readFileSync(
  new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json', import.meta.url),
  'utf8',
))) as NativeWorkbook

describe('native XLSX round trip', () => {
  it('ships the meaningful repository-owned launch workbook', () => {
    const bytes = readFileSync(new URL('../public/native-fixture/launch-readiness-plan.xlsx', import.meta.url))
    expect(bytes.subarray(0, 4).toString('hex')).toBe('504b0304')
    expect(bytes.byteLength).toBeGreaterThan(4_000)
  })

  it('selects safe literal targets and builds the joined native transaction', () => {
    const targets = editableTargets(workbook)
    expect(targets.length).toBeGreaterThan(0)
    expect(displayCellValue(targets[0])).toBeTypeOf('string')
    expect(buildCellMutation(workbook, targets[0]!, 'after', 'demo-edit')).toEqual({
      protocol: 'injoffice.xlsx.mutations',
      version: 1,
      batch_id: 'demo-edit',
      expected_revision: workbook.source.package_sha256,
      operations: [{
        operation_id: 'demo-edit',
        sheet_id: targets[0]!.sheetId,
        kind: 'cell.set_value',
        cell: { row: targets[0]!.row, column: targets[0]!.column },
        value: 'after',
      }],
    })
  })

  it('checks the protocol boundary and creates a safe output filename', () => {
    expect(decodeNativeWorkbook(workbook)).toBe(workbook)
    expect(() => decodeNativeWorkbook({ ...workbook, protocol: 'something-else' })).toThrow(/protocol/i)
    expect(() => decodeNativeWorkbook({ ...workbook, sheets: [{ ...workbook.sheets[0], cells: [{}] }] })).toThrow(/cell/i)
    expect(downloadName('Quarterly plan.xlsx')).toBe('Quarterly-plan-injoffice.xlsx')
  })
})
