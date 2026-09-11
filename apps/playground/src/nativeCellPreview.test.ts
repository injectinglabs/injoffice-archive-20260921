import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { nativeCellPreview } from './nativeCellPreview'
import { decodeNativeWorkbook, displayCellValue, type NativeCell, type NativeWorkbook } from './nativeRoundTrip'

const original = decodeNativeWorkbook(JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json', import.meta.url), 'utf8')))
function sample(format: string, lexical: string, date1904?: boolean): { workbook: NativeWorkbook; cell: NativeCell } {
  const style = original.styles[0]!
  return {
    workbook: { ...original, date1904, styles: [{ ...style, effective: { ...style.effective, number_format: format, unsupported: [] } }] },
    cell: { row: 0, column: 0, ref: 'A1', style_id: 0, editable: true, value: { kind: 'number', storage: 'number', lexical, rich: false } },
  }
}

describe('native spreadsheet display integration', () => {
  it.each([
    ['0.00', '12.345', '12.35'],
    ['0.00%', '0.125', '12.50%'],
    ['"$"#,##0.00', '1234.56', '$1,234.56'],
    ['0.00" €"', '1234.56', '1234.56 €'],
    ['General', '9007199254740993', '9007199254740993'],
  ])('uses the native formatter for %s', (format, lexical, expected) => {
    const { workbook, cell } = sample(format, lexical)
    const before = JSON.stringify({ workbook, cell })
    expect(nativeCellPreview(workbook, cell)).toEqual({ text: expected, cached: false })
    expect(displayCellValue(cell)).toBe(lexical)
    expect(JSON.stringify({ workbook, cell })).toBe(before)
  })

  it('uses the explicit workbook date system and reports unavailable date evidence', () => {
    for (const [date1904, lexical, expected] of [[false, '1', '1900-01-01'], [true, '0', '1904-01-01']] as const) {
      const { workbook, cell } = sample('yyyy-mm-dd', lexical, date1904)
      expect(nativeCellPreview(workbook, cell).text).toBe(expected)
    }
    const { workbook, cell } = sample('yyyy-mm-dd', '1')
    expect(nativeCellPreview(workbook, cell)).toMatchObject({ text: '1', warning: expect.any(String) })
  })

  it('uses saved formula results without granting editability or evaluating formulas', () => {
    const { workbook, cell } = sample('0.00%', '0.125')
    const formula: NativeCell = { ...cell, editable: false, value: undefined, formula: { type: 'normal', text: '1/8', cached: cell.value } }
    expect(nativeCellPreview(workbook, formula)).toEqual({ text: '12.50%', cached: true })
    expect(displayCellValue(formula)).toBe('=1/8')
    expect(formula.editable).toBe(false)
    expect(nativeCellPreview(workbook, { ...formula, formula: { type: 'normal', text: '1/8' } })).toMatchObject({ text: '—', warning: expect.stringContaining('No saved formula result') })
  })

  it('marks raw fallbacks instead of guessing unsupported currency/accounting formats', () => {
    for (const format of ['$#,##0.00', '[Red]0.00', '0.00;[Red]-0.00']) {
      const { workbook, cell } = sample(format, '1234.56')
      expect(nativeCellPreview(workbook, cell)).toMatchObject({ text: '1234.56', warning: expect.stringContaining('Unsupported') })
    }
  })

  it('marks missing style or number format evidence instead of assuming General', () => {
    const { workbook, cell } = sample('0.00%', '0.125')
    const style = workbook.styles[0]!
    for (const styles of [[], [{ ...style, effective: { ...style.effective, number_format: undefined } }], [{ ...style, effective: { ...style.effective, unsupported: ['number-format' as const] } }]]) {
      expect(nativeCellPreview({ ...workbook, styles }, cell)).toEqual({ text: '0.125', cached: false, warning: 'Number format unavailable; showing the stored value.' })
      expect(displayCellValue(cell)).toBe('0.125')
    }
  })

  it('displays booleans, cached errors, and blank cells', () => {
    const { workbook, cell } = sample('General', '1')
    expect(nativeCellPreview(workbook, { ...cell, value: { kind: 'boolean', storage: 'boolean', lexical: '1', rich: false } }).text).toBe('TRUE')
    expect(nativeCellPreview(workbook, { ...cell, formula: { type: 'normal', text: '1/0', cached: { kind: 'error', storage: 'error', lexical: '#DIV/0!', rich: false } } }).text).toBe('#DIV/0!')
    expect(nativeCellPreview(workbook, undefined)).toEqual({ text: '' })
  })
})
