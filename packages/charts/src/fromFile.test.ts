import { describe, expect, it } from 'vitest'
import { colToIndex, parseRef, specFromFileChart, specsFromFileCharts } from './fromFile'
import type { FileChartInfo } from './fromFile'

const info = (over: Partial<FileChartInfo> = {}): FileChartInfo => ({
  part: 'xl/charts/chart1.xml',
  type: 'column',
  title: 'Revenue by Quarter',
  series: [
    { name: 'Revenue', nameRef: 'Data!$B$1', categoriesRef: 'Data!$A$2:$A$5', valuesRef: 'Data!$B$2:$B$5' },
    { name: 'Costs', nameRef: 'Data!$C$1', categoriesRef: 'Data!$A$2:$A$5', valuesRef: 'Data!$C$2:$C$5' },
  ],
  ...over,
})

const sheets = { Data: 'sheet-id-1' }

describe('colToIndex', () => {
  it('maps single and multi letter columns', () => {
    expect(colToIndex('A')).toBe(0)
    expect(colToIndex('Z')).toBe(25)
    expect(colToIndex('AA')).toBe(26)
    expect(colToIndex('AZ')).toBe(51)
    expect(colToIndex('BA')).toBe(52)
  })
})

describe('parseRef', () => {
  it('parses absolute range refs', () => {
    expect(parseRef('Data!$B$2:$B$5')).toEqual({
      sheetName: 'Data', startRow: 1, startColumn: 1, endRow: 4, endColumn: 1,
    })
  })
  it('parses single cells and relative refs', () => {
    expect(parseRef('Data!C1')).toEqual({
      sheetName: 'Data', startRow: 0, startColumn: 2, endRow: 0, endColumn: 2,
    })
  })
  it('parses quoted sheet names with escapes', () => {
    expect(parseRef("'It''s 2026'!$A$1")?.sheetName).toBe("It's 2026")
  })
  it('rejects garbage', () => {
    expect(parseRef('$A$1')).toBeNull()
    expect(parseRef('Data!')).toBeNull()
    expect(parseRef('Data!B5:B2')).toBeNull()
  })
})

describe('specFromFileChart', () => {
  it('builds a spec over the bounding block of all series refs', () => {
    const conv = specFromFileChart(info(), sheets)
    expect(conv).not.toBeNull()
    // A1:C5 — headers row 1 through data row 5, columns A..C.
    expect(conv!.spec.range).toEqual({
      sheetId: 'sheet-id-1', startRow: 0, startColumn: 0, endRow: 4, endColumn: 2,
    })
    expect(conv!.spec.type).toBe('column')
    expect(conv!.spec.title).toBe('Revenue by Quarter')
    expect(conv!.spec.id).toBe('filechart-xl-charts-chart1-xml')
  })

  it('carries the file anchor through', () => {
    const anchor = { FromCol: 4, FromRow: 1, ToCol: 12, ToRow: 16 }
    const conv = specFromFileChart(info(), sheets, anchor)
    expect(conv!.cellAnchor).toEqual(anchor)
  })

  it('carries a valid native identity into the mounted spec', () => {
    const identity = { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 }
    expect(specFromFileChart(info({ identity }), sheets)?.spec.nativeIdentity).toEqual(identity)
    expect(specFromFileChart(info({ identity: { ...identity, part: 'xl/charts/chart2.xml' } }), sheets)?.spec.nativeIdentity).toBeUndefined()
  })

  it('returns null for unrenderable types and unknown sheets', () => {
    expect(specFromFileChart(info({ type: 'unknown' }), sheets)).toBeNull()
    expect(specFromFileChart(info(), { Other: 'x' })).toBeNull()
    expect(specFromFileChart(info({ series: [{ name: 'n' }] }), sheets)).toBeNull()
  })
})

describe('specsFromFileCharts', () => {
  it('converts what it can and counts what it cannot', () => {
    const { conversions, skipped } = specsFromFileCharts(
      [info(), info({ part: 'xl/charts/chart2.xml', type: 'unknown' })],
      { 'xl/charts/chart1.xml': { FromCol: 1, FromRow: 1, ToCol: 8, ToRow: 12 } },
      sheets,
    )
    expect(conversions).toHaveLength(1)
    expect(conversions[0].cellAnchor?.ToCol).toBe(8)
    expect(skipped).toBe(1)
  })
})
