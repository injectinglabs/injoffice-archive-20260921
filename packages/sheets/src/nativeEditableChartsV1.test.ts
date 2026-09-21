import { describe, expect, it } from 'vitest'
import { decodeNativeEditableChartsV1 } from './nativeEditableChartsV1.js'

const chart = {
  identity: { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 1 },
  sheet_id: '1',
  fingerprint_sha256: `sha256:${'a'.repeat(64)}`,
  chart_type: 'bar',
  title: 'Revenue',
  range: { row: 0, column: 0, end_row: 4, end_column: 2 },
  anchor: { from_row: 0, from_column: 4, to_row: 18, to_column: 12 },
  editable: true,
  categories: ['Q1', 'Q2'],
  series: [{ name: 'North', values: ['10', '20'] }],
}

describe('native editable charts v1', () => {
  it('owns a copy of a qualified chart', () => {
    const decoded = decodeNativeEditableChartsV1([chart])
    expect(decoded[0]).toEqual(chart)
    decoded[0]!.categories[0] = 'changed'
    expect(chart.categories[0]).toBe('Q1')
  })

  it('refuses an editable chart without series and an unknown field', () => {
    expect(() => decodeNativeEditableChartsV1([{ ...chart, series: [] }])).toThrow()
    expect(() => decodeNativeEditableChartsV1([{ ...chart, extra: true }])).toThrow()
  })
})
