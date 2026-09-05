import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { projectNativeWorkbookV1, validateNativeWorkbookV1 } from './index.js'
import type { NativeWorkbookV1 } from './index.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const fixturePath = resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v1/valid/lexical-render.json')

function mergedWorkbook(): NativeWorkbookV1 {
  const workbook = JSON.parse(readFileSync(fixturePath, 'utf8')) as NativeWorkbookV1
  const sheet = workbook.sheets[0] as unknown as {
    rows: unknown[]
    columns: unknown[]
    cells: unknown[]
    merged_ranges: unknown[]
  }
  sheet.rows.push({ row: 1, hidden: true, custom_height: false })
  sheet.columns.push({ column: 4, end_column: 4, hidden: true, custom_width: false, best_fit: false, style_id: 1 })
  sheet.cells.push(
    { row: 1, column: 3, ref: 'D2', style_id: 1, editable: false },
    { row: 1, column: 4, ref: 'E2', style_id: 0, editable: false },
  )
  sheet.merged_ranges = [{ ref: 'D2:F3', row: 1, column: 3, end_row: 2, end_column: 5, editable: false }]
  const location = ['MERGED_CELLS', 'merges', `sheet:${workbook.sheets[0].id}`, workbook.sheets[0].part_name, '', ''].join('\0')
  ;(workbook as unknown as { unsupported: unknown[] }).unsupported.push({
    id: `unsupported:${createHash('sha256').update(location).digest('hex')}`,
    code: 'MERGED_CELLS',
    capability: 'merges',
    scope_id: `sheet:${workbook.sheets[0].id}`,
    part_name: workbook.sheets[0].part_name,
    preservation: 'preserve-exact',
    message: 'merged ranges are projected exactly while source SpreadsheetML remains authoritative',
  })
  return workbook
}

describe('native XLSX merged-range projection', () => {
  it('exposes an exact renderer-neutral span with top-left content authority and no writer authority', () => {
    const model = projectNativeWorkbookV1(mergedWorkbook())
    expect(model.sheets[0].merged_ranges).toEqual([{
      ref: 'D2:F3',
      row: 1,
      column: 3,
      end_row: 2,
      end_column: 5,
      row_span: 2,
      column_span: 3,
      top_left: { row: 1, column: 3, ref: 'D2' },
      content_authority: 'top-left',
      covered_cell_content: 'blank',
      editable: false,
      source_part: 'Worksheets/Sheet1.xml',
      source_revision: mergedWorkbook().revision,
    }])
    expect(model.sheets[0].mutation_authority.editable).toBe(true)
  })

  it('does not renumber merges across hidden dimensions or synthesize shared styles', () => {
    const model = projectNativeWorkbookV1(mergedWorkbook())
    const sheet = model.sheets[0]
    expect(sheet.rows.at(-1)).toEqual({ row: 1, hidden: true, custom_height: false })
    expect(sheet.columns.at(-1)).toEqual({ column: 4, end_column: 4, hidden: true, custom_width: false, best_fit: false, style_id: 1 })
    expect(sheet.merged_ranges[0]).toMatchObject({ row: 1, column: 3, end_row: 2, end_column: 5 })
    expect(sheet.cells.slice(-2)).toEqual([
      { row: 1, column: 3, ref: 'D2', style_id: 1, content: { kind: 'blank' }, editable: false },
      { row: 1, column: 4, ref: 'E2', style_id: 0, content: { kind: 'blank' }, editable: false },
    ])
  })

  it('refuses ambiguous covered values and requires every modeled merged cell to be read-only', () => {
    const coveredValue = mergedWorkbook()
    ;(coveredValue.sheets[0].cells.at(-1) as unknown as { value: unknown }).value = {
      kind: 'number', storage: 'number', lexical: '7', rich: false,
    }
    const coveredValueResult = validateNativeWorkbookV1(coveredValue)
    expect(coveredValueResult.ok).toBe(false)
    if (!coveredValueResult.ok) expect(coveredValueResult.issues.map((issue) => issue.message)).toContain('covered non-anchor merged cell must be blank; a value or formula is ambiguous')

    const editableAnchor = mergedWorkbook()
    ;(editableAnchor.sheets[0].cells.at(-2) as unknown as { editable: boolean }).editable = true
    const editableAnchorResult = validateNativeWorkbookV1(editableAnchor)
    expect(editableAnchorResult.ok).toBe(false)
    if (!editableAnchorResult.ok) expect(editableAnchorResult.issues.some((issue) => issue.message.includes('must be mutation-refused'))).toBe(true)
  })

  it('keeps a maximum-grid merge sparse instead of expanding covered cells', () => {
    const workbook = mergedWorkbook()
    ;(workbook.sheets[0].merged_ranges[0] as unknown as Record<string, unknown>).ref = 'A2:XFD1048576'
    ;(workbook.sheets[0].merged_ranges[0] as unknown as Record<string, unknown>).column = 0
    ;(workbook.sheets[0].merged_ranges[0] as unknown as Record<string, unknown>).end_row = 1048575
    ;(workbook.sheets[0].merged_ranges[0] as unknown as Record<string, unknown>).end_column = 16383
    const model = projectNativeWorkbookV1(workbook)
    expect(model.sheets[0].merged_ranges[0]).toMatchObject({ row_span: 1048575, column_span: 16384 })
    expect(model.sheets[0].cells).toHaveLength(8)
  })
})
