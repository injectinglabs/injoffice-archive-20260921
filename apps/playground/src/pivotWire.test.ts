import { describe, expect, it } from 'vitest'
import { playgroundPivotSpec, playgroundPivotWire } from './pivotWire'

const header = ['Region', 'Owner', 'Quarter', 'Revenue', 'Units']

describe('playground pivot native wire', () => {
  it('converts a baked pivot spec and reports native representability', () => {
    const spec = playgroundPivotSpec({
      rows: ['Region'],
      columns: ['Quarter'],
      valueField: 'Revenue',
      aggregation: 'sum',
      rowCount: 7,
      columnCount: 5,
    })
    const result = playgroundPivotWire(spec, header, { Region: ['West', 'East', 'North'], Quarter: ['Q1', 'Q2'] })
    expect(result.representability.issues.filter((issue) => issue.severity === 'unsupported')).toEqual([])
    expect(result.wire.pivots[0]).toMatchObject({
      sourceSheetName: 'Sales',
      targetSheetName: 'Sales',
      rowFields: ['Region'],
      colFields: ['Quarter'],
    })
    expect(result.wire.skipped).toEqual([])
  })
})
