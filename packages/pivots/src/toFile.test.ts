import { describe, expect, it } from 'vitest'
import { toWirePivotRemove, toWirePivots, toWirePivotUpdate } from './toFile'
import type { PivotSpec } from './types'

const spec = (over: Partial<PivotSpec> = {}): PivotSpec => ({
  id: 'pivot-1a',
  source: { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 4, endColumn: 3 },
  rows: ['Quarter'],
  columns: [],
  values: [{ field: 'Profit', agg: 'sum' }],
  target: { sheetId: 's1', startRow: 0, startColumn: 5 },
  ...over,
})

const ctx = {
  sheetNameOf: (id: string) => (id === 's1' ? 'Data' : null),
  headerRowOf: () => ['Quarter', 'Revenue', 'Costs', 'Profit'],
}

describe('toWirePivots', () => {
  it('maps a spec to the gateway wire shape', () => {
    const { pivots, skipped } = toWirePivots([spec()], ctx)
    expect(skipped).toEqual([])
    expect(pivots).toEqual([
      {
        sourceSheetName: 'Data',
        sourceRef: 'A1:D5',
        fields: ['Quarter', 'Revenue', 'Costs', 'Profit'],
        targetSheetName: 'Data',
        targetCellRef: 'F1',
        rowFields: ['Quarter'],
        colFields: undefined,
        dataFields: [{ field: 'Profit', agg: 'sum' }],
        name: 'InjOffice_pivot_1a',
      },
    ])
  })

  it('labels blank header cells and skips broken specs with reasons', () => {
    const { pivots } = toWirePivots([spec()], { ...ctx, headerRowOf: () => ['Quarter', '', null, 'Profit'] })
    expect(pivots[0].fields).toEqual(['Quarter', 'Column 2', 'Column 3', 'Profit'])

    const { pivots: none, skipped } = toWirePivots(
      [spec({ source: { ...spec().source, sheetId: 'gone' } }), spec({ id: 'novals', values: [] })],
      ctx,
    )
    expect(none).toEqual([])
    expect(skipped).toHaveLength(2)
    expect(skipped[0]).toContain('sheet no longer exists')
    expect(skipped[1]).toContain('no value fields')
  })

  it('builds identity-bound native update and remove requests', () => {
    const hydrated = spec({ nativeIdentity: { part: 'xl/pivotTables/pivotTable7.xml' } })
    expect(toWirePivotRemove(hydrated)).toEqual({
      request: {
        operation: 'remove',
        identity: { part: 'xl/pivotTables/pivotTable7.xml' },
      },
      skipped: [],
    })
    expect(toWirePivotUpdate(hydrated, ctx)).toEqual({
      request: {
        operation: 'update',
        identity: { part: 'xl/pivotTables/pivotTable7.xml' },
        pivot: {
          sourceSheetName: 'Data',
          sourceRef: 'A1:D5',
          fields: ['Quarter', 'Revenue', 'Costs', 'Profit'],
          targetSheetName: 'Data',
          targetCellRef: 'F1',
          rowFields: ['Quarter'],
          colFields: undefined,
          dataFields: [{ field: 'Profit', agg: 'sum' }],
          name: 'InjOffice_pivot_1a',
        },
      },
      skipped: [],
    })
  })

  it('maps bounded native page, member-filter, sort and grand-total state', () => {
    const configured = spec({
      pageFields: [{ field: 'Costs', selectedItem: 'Fixed' }],
      memberFilters: [{ field: 'Quarter', mode: 'include', values: ['Q2'] }],
      sorts: [{ field: 'Quarter', direction: 'descending' }],
      rows: ['Quarter'],
      grandTotals: false,
    })
    const withMembers = {
      ...ctx,
      fieldMembersOf: (_source: PivotSpec['source'], field: string) => ({
        Costs: [{ value: 'Fixed', kind: 'string' as const }, { value: 'Variable', kind: 'string' as const }],
        Quarter: [{ value: 'Q1', kind: 'string' as const }, { value: 'Q2', kind: 'string' as const }],
      })[field as 'Costs' | 'Quarter'] ?? null,
    }
    const result = toWirePivots([configured], withMembers)
    expect(result.skipped).toEqual([])
    expect(result.pivots[0]).toMatchObject({
      pageFields: [{ field: 'Costs', selectedItem: 'Fixed' }],
      memberFilters: [{ field: 'Quarter', excludedItems: ['Q1'] }],
      sorts: [{ field: 'Quarter', direction: 'descending' }],
      fieldMembers: [
        { field: 'Quarter', items: [{ value: 'Q1', kind: 'string' }, { value: 'Q2', kind: 'string' }] },
        { field: 'Costs', items: [{ value: 'Fixed', kind: 'string' }, { value: 'Variable', kind: 'string' }] },
      ],
      grandTotals: false,
    })
  })

  it('refuses stateful conversion without unambiguous complete members', () => {
    const configured = spec({ memberFilters: [{ field: 'Quarter', mode: 'exclude', values: ['Q1'] }] })
    expect(toWirePivots([configured], ctx).skipped[0]).toContain('MEMBER_INVENTORY_REQUIRED')
    const ambiguous = {
      ...ctx,
      fieldMembersOf: () => [{ value: 'Q1', kind: 'string' as const }, { value: 'Q1', kind: 'number' as const }],
    }
    expect(toWirePivots([configured], ambiguous).skipped[0]).toContain('DUPLICATE_MEMBER')
  })

  it('refuses lifecycle requests that lack native identity', () => {
    expect(toWirePivotRemove(spec()).request).toBeUndefined()
    expect(toWirePivotRemove(spec()).skipped[0]).toContain('no hydrated native identity')
    expect(toWirePivotUpdate(spec(), ctx).request).toBeUndefined()
  })
})
