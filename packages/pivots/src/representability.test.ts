import { describe, expect, it } from 'vitest'
import { assessPivotRepresentability } from './representability'
import type { PivotSpec } from './types'

const configured: PivotSpec = {
  id: 'p',
  source: { sheetId: 's', startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
  rows: ['Region'],
  columns: [],
  values: [{ field: 'Sales', agg: 'sum' }],
  pageFields: [{ field: 'Quarter', selectedItem: 'Q1' }],
  sorts: [{ field: 'Region', direction: 'ascending' }],
  target: { sheetId: 's', startRow: 0, startColumn: 4 },
}

describe('assessPivotRepresentability', () => {
  it('accepts the bounded page/filter/label-sort subset with complete members', () => {
    const result = assessPivotRepresentability(configured, {
      fields: ['Region', 'Quarter', 'Sales'],
      membersOf: () => [{ value: 'Q1', kind: 'string' }],
    })
    expect(result).toEqual({ representable: true, issues: [] })
  })

  it('returns typed refusals for value-sort placement and ambiguous members', () => {
    const result = assessPivotRepresentability({
      ...configured,
      sorts: [{ field: 'Sales', direction: 'descending' }],
    }, {
      fields: ['Region', 'Quarter', 'Sales'],
      membersOf: () => [{ value: 'Q1', kind: 'string' }, { value: 'Q1', kind: 'number' }],
    })
    expect(result.representable).toBe(false)
    expect(result.issues.map(({ code }) => code)).toEqual(expect.arrayContaining(['SORT_AXIS_UNSUPPORTED', 'DUPLICATE_MEMBER']))
  })
})
