import { describe, expect, it } from 'vitest'
import { XLSX_MAX_OUTLINE_DEPTH, createNativeXlsxOutlineWrite, decodeNativeXlsxOutlineSnapshot } from './native'
import type { OutlineGroup } from './types'

const group = (id: string, start: number, end: number, axis: 'row' | 'column' = 'row'): OutlineGroup => ({
  id,
  sheetId: '7',
  axis,
  start,
  end,
  collapsed: false,
})

describe('native XLSX outline wire helpers', () => {
  it('decodes exact native bands and groups', () => {
    const result = decodeNativeXlsxOutlineSnapshot({
      sheet_id: '7',
      summary_below: true,
      summary_right: false,
      rows: [{ start: 1, end: 2, outline_level: 1, hidden: true, collapsed: false }],
      columns: [{ start: 3, end: 5, outline_level: 1, hidden: false, collapsed: false }],
      groups: [
        { id: 'rows', sheet_id: '7', axis: 'row', start: 1, end: 2, collapsed: true },
        { id: 'columns', sheet_id: '7', axis: 'column', start: 3, end: 5, collapsed: false },
      ],
    })
    expect(result).toEqual({
      ok: true,
      value: {
        sheetId: '7',
        summaryBelow: true,
        summaryRight: false,
        rows: [{ start: 1, end: 2, outlineLevel: 1, hidden: true, collapsed: false }],
        columns: [{ start: 3, end: 5, outlineLevel: 1, hidden: false, collapsed: false }],
        groups: [
          { id: 'rows', sheetId: '7', axis: 'row', start: 1, end: 2, collapsed: true },
          { id: 'columns', sheetId: '7', axis: 'column', start: 3, end: 5, collapsed: false },
        ],
      },
    })
  })

  it('encodes the Go write contract with SpreadsheetML defaults', () => {
    const result = createNativeXlsxOutlineWrite('7', [group('rows', 1, 4)])
    expect(result).toEqual({
      ok: true,
      value: {
        sheet_id: '7',
        summary_below: true,
        summary_right: true,
        groups: [{ id: 'rows', sheet_id: '7', axis: 'row', start: 1, end: 4, collapsed: false }],
      },
    })
  })

  it('rejects unknown fields, malformed bands, and sheet mismatches', () => {
    const malformed = decodeNativeXlsxOutlineSnapshot({
      sheet_id: '7',
      summary_below: true,
      summary_right: true,
      rows: [
        { start: 2, end: 3, outline_level: 1, hidden: false, collapsed: false },
        { start: 3, end: 4, outline_level: 1, hidden: false, collapsed: false },
      ],
      columns: [],
      groups: [{ id: 'bad', sheet_id: '9', axis: 'row', start: 0, end: 1, collapsed: false }],
      extra: true,
    })
    expect(malformed.ok).toBe(false)
    if (!malformed.ok) {
      expect(malformed.issues.join('\n')).toMatch(/not supported/)
      expect(malformed.issues.join('\n')).toMatch(/overlaps/)
      expect(malformed.issues.join('\n')).toMatch(/must match/)
    }
  })

  it('rejects an eighth nested group at the native persistence boundary', () => {
    const groups = Array.from({ length: XLSX_MAX_OUTLINE_DEPTH + 1 }, (_, depth) => group(`depth-${depth}`, depth, 20 - depth))
    const result = createNativeXlsxOutlineWrite('7', groups)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.some((issue) => issue.code === 'MAX_DEPTH')).toBe(true)
  })

  it('rejects crossing writes before they reach native code', () => {
    const result = createNativeXlsxOutlineWrite('7', [group('left', 1, 5), group('right', 3, 7)])
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.some((issue) => issue.code === 'CROSSING')).toBe(true)
  })
})
