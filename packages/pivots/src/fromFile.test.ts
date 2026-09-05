import { describe, expect, it } from 'vitest'
import { pivotsFromFile, type FilePivotInfo } from './fromFile'

const nativePivot: FilePivotInfo = {
  part: 'xl/pivotTables/pivotTable1.xml',
  name: 'Profit Pivot',
  cacheId: 1,
  sourceSheetName: 'Data',
  sourceRef: '$A$1:$D$5',
  fields: ['Quarter', 'Revenue', 'Costs', 'Profit'],
  targetSheetName: 'Summary',
  targetRef: 'F2:J20',
  rowFields: ['Quarter'],
  colFields: [],
  dataFields: [{ field: 'Profit', agg: 'sum' }],
  grandTotals: true,
}

describe('pivotsFromFile', () => {
  it('hydrates a native pivot into an editable spec', () => {
    const result = pivotsFromFile([nativePivot], {
      sheetIdOf: (name) => ({ Data: 'sheet-data', Summary: 'sheet-summary' })[name] ?? null,
    })
    expect(result.skipped).toEqual([])
    expect(result.pivots).toEqual([{
      id: 'Profit_Pivot',
      nativeIdentity: { part: 'xl/pivotTables/pivotTable1.xml' },
      source: { sheetId: 'sheet-data', startRow: 0, startColumn: 0, endRow: 4, endColumn: 3 },
      rows: ['Quarter'],
      columns: [],
      values: [{ field: 'Profit', agg: 'sum' }],
      target: { sheetId: 'sheet-summary', startRow: 1, startColumn: 5 },
      grandTotals: true,
    }])
  })

  it('refuses warnings and unsupported aggregations instead of losing state', () => {
    const warned = { ...nativePivot, warnings: ['page fields are not supported'] }
    const median = { ...nativePivot, name: 'Median', dataFields: [{ field: 'Profit', agg: 'median' }] }
    const result = pivotsFromFile([warned, median], { sheetIdOf: () => 'sheet' })
    expect(result.pivots).toEqual([])
    expect(result.skipped).toHaveLength(2)
  })

  it('hydrates supported native page, exclusion and label-sort state', () => {
    const configured: FilePivotInfo = {
      ...nativePivot,
      pageFields: [{ field: 'Costs', selectedItem: 'Fixed' }],
      memberFilters: [{ field: 'Quarter', excludedItems: ['Q1'] }],
      sorts: [{ field: 'Quarter', direction: 'descending' }],
    }
    const result = pivotsFromFile([configured], { sheetIdOf: (name) => name })
    expect(result.skipped).toEqual([])
    expect(result.pivots[0]).toMatchObject({
      pageFields: [{ field: 'Costs', selectedItem: 'Fixed' }],
      memberFilters: [{ field: 'Quarter', mode: 'exclude', values: ['Q1'] }],
      sorts: [{ field: 'Quarter', direction: 'descending' }],
    })
  })
})
