import { describe, expect, it } from 'vitest'
import { WORKBOOK_MUTATION_PROTOCOL, decodeWorkbookMutationBatch } from './mutationProtocol.js'

const range = { row: 0, column: 0, end_row: 4, end_column: 2 }
const anchor = { from_row: 0, from_column: 4, to_row: 18, to_column: 12 }
const identity = { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 1 }
const fingerprint = `sha256:${'a'.repeat(64)}`

const batch = (operations: unknown[]) => ({
  protocol: WORKBOOK_MUTATION_PROTOCOL,
  version: 1,
  batch_id: 'chart-1',
  expected_revision: 'sha256:workbook',
  operations,
})

const base = (kind: string, extra: Record<string, unknown>) => ({
  operation_id: kind.replace('.', '-'),
  kind,
  sheet_id: '1',
  ...extra,
})

describe('chart mutation protocol v1', () => {
  it('accepts insert, update, and delete in listed order', () => {
    const input = batch([
      base('chart.insert', { chart_type: 'bar', title: 'Revenue', range, anchor }),
      base('chart.update', { identity, expected_fingerprint_sha256: fingerprint, chart_type: 'line', title: 'Revenue', range, anchor }),
      base('chart.delete', { identity, expected_fingerprint_sha256: fingerprint }),
    ])
    const decoded = decodeWorkbookMutationBatch(input)
    expect(decoded.ok).toBe(true)
    if (!decoded.ok) return
    expect(decoded.value.operations.map((operation) => operation.kind)).toEqual(['chart.insert', 'chart.update', 'chart.delete'])
    expect(decoded.value.operations[0]).toEqual({
      operation_id: 'chart-insert',
      sheet_id: '1',
      kind: 'chart.insert',
      chart_type: 'bar',
      title: 'Revenue',
      range,
      anchor,
    })
  })

  it('accepts pie and an empty title', () => {
    const decoded = decodeWorkbookMutationBatch(batch([
      base('chart.insert', { chart_type: 'pie', title: '', range, anchor }),
    ]))
    expect(decoded.ok).toBe(true)
    if (decoded.ok) expect(decoded.value.operations[0]).toMatchObject({ chart_type: 'pie', title: '' })
  })

  it('refuses a single-cell source, heatmap type, and identity on insert', () => {
    const single = decodeWorkbookMutationBatch(batch([
      base('chart.insert', { chart_type: 'bar', title: '', range: { row: 0, column: 0, end_row: 0, end_column: 0 }, anchor }),
    ]))
    expect(single.ok).toBe(false)
    if (!single.ok) expect(single.issues).toEqual([expect.objectContaining({ code: 'INVALID_RANGE', path: '/operations/0/range' })])

    const heatmap = decodeWorkbookMutationBatch(batch([
      base('chart.insert', { chart_type: 'heatmap', title: '', range, anchor }),
    ]))
    expect(heatmap.ok).toBe(false)
    if (!heatmap.ok) expect(heatmap.issues).toEqual([expect.objectContaining({ code: 'INVALID_VALUE', path: '/operations/0/chart_type' })])

    const extra = decodeWorkbookMutationBatch(batch([
      base('chart.insert', { chart_type: 'bar', title: '', range, anchor, identity }),
    ]))
    expect(extra.ok).toBe(false)
    if (!extra.ok) expect(extra.issues).toEqual([expect.objectContaining({ code: 'UNKNOWN_FIELD', path: '/operations/0/identity' })])
  })

  it('requires a sha256 fingerprint and native chart part paths on update', () => {
    const decoded = decodeWorkbookMutationBatch(batch([
      base('chart.update', {
        identity: { part: 'charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 1 },
        expected_fingerprint_sha256: 'not-a-digest',
        chart_type: 'column',
        title: '',
        range,
        anchor,
      }),
    ]))
    expect(decoded.ok).toBe(false)
    if (!decoded.ok) {
      expect(decoded.issues).toEqual(expect.arrayContaining([
        expect.objectContaining({ path: '/operations/0/identity/part' }),
        expect.objectContaining({ path: '/operations/0/expected_fingerprint_sha256' }),
      ]))
    }
  })
})
