import { describe, expect, it } from 'vitest'
import type { NativeDocxTableCellV1, NativeDocxTableV1 } from '../../../packages/docs/src/nativeContract'
import { nativeDocxHighlight, nativeDocxTableRows } from './docsNativePreview'

const anchor = { part_name: 'word/document.xml', path: '/p', start_byte: 1, end_byte: 2, xml_sha256: `sha256:${'a'.repeat(64)}` }
const edit_policy = { mode: 'read-only' as const, allowed_operations: [] }
function cell(id: string, merge: NativeDocxTableCellV1['vertical_merge'] = 'none', span = 1, text = ''): NativeDocxTableCellV1 {
  return { id, anchor, vertical_merge: merge, grid_span: span, paragraphs: [{ id: `${id}:p`, anchor, edit_policy, properties: {}, runs: text ? [{ id: `${id}:r`, anchor, kind: 'text', text }] : [] }] }
}
function table(rows: NativeDocxTableCellV1[][]): NativeDocxTableV1 {
  return { id: 'table', anchor, edit_policy, rows: rows.map((cells, index) => ({ id: `row:${index}`, anchor, repeat_header: false, cells })) }
}

describe('approximate DOCX table projection', () => {
  it('joins equal grid intervals across consecutive rows, without mutating the source', () => {
    const source = table([[cell('a', 'restart', 2, 'Heading'), cell('b')], [cell('c', 'continue', 2), cell('d')], [cell('e', 'continue', 2), cell('f')]])
    const before = JSON.stringify(source)
    const rows = nativeDocxTableRows(source)
    expect(rows.map((row) => row.map((entry) => [entry.cell.id, entry.column, entry.rowSpan]))).toEqual([
      [['a', 0, 3], ['b', 2, 1]], [['d', 2, 1]], [['f', 2, 1]],
    ])
    expect(JSON.stringify(source)).toBe(before)
  })
  it('does not join mismatched spans or merge through an intervening row', () => {
    const rows = nativeDocxTableRows(table([[cell('a', 'restart', 2)], [cell('b', 'continue')], [cell('c', 'continue', 2)]]))
    expect(rows[0][0].rowSpan).toBe(1)
    expect(rows[1][0].orphanContinuation).toBe(true)
    expect(rows[2][0].orphanContinuation).toBe(true)
  })
  it('preserves unexpected text in continuation cells instead of silently hiding it', () => {
    const rows = nativeDocxTableRows(table([[cell('a', 'restart')], [cell('b', 'continue', 1, 'Must remain visible')]]))
    expect(rows[0][0].rowSpan).toBe(1)
    expect(rows[1][0].cell.paragraphs[0].runs[0].text).toBe('Must remain visible')
    expect(rows[1][0].orphanContinuation).toBe(true)
  })
  it('does not mistake same-width cells in different grid columns for continuations', () => {
    const rows = nativeDocxTableRows(table([[cell('a'), cell('b', 'restart')], [cell('c', 'continue'), cell('d')]]))
    expect(rows[0][1].rowSpan).toBe(1)
    expect(rows[1][0].orphanContinuation).toBe(true)
  })
})

describe('OOXML highlight projection', () => {
  it('maps named highlights explicitly and ignores unsupported tokens', () => {
    expect(nativeDocxHighlight('darkYellow')).toBe('#808000')
    expect(nativeDocxHighlight('yellow')).toBe('#ffff00')
    expect(nativeDocxHighlight('none')).toBeUndefined()
    expect(nativeDocxHighlight('url(https://example.com)')).toBeUndefined()
    expect(nativeDocxHighlight('constructor')).toBeUndefined()
  })
})
