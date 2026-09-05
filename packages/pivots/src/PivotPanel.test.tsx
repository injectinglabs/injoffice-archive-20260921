import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it } from 'vitest'
import { PivotCommandController, type PivotUndoRecord } from './commands'
import { createPivotPanelCommandBindings } from './PivotPanel'
import { PivotManager } from './manager'
import type { PivotSpec } from './types'

const GRID = [
  ['Region', 'Product', 'Quarter', 'Sales'],
  ['West', 'Widget', 'Q1', 20],
  ['East', 'Gadget', 'Q2', 35],
]

function manager(): PivotManager {
  const sheet = {
    getSheetId: () => 'sheet-1',
    getRange: () => ({ getValues: () => GRID, setValues: () => undefined }),
  }
  const workbook = {
    getActiveSheet: () => sheet,
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }
  return new PivotManager({ getActiveWorkbook: () => workbook } as unknown as FUniver)
}

const SPEC: PivotSpec = {
  id: 'pivot-1',
  source: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 3 },
  rows: ['Region'],
  columns: [],
  values: [{ field: 'Sales', agg: 'sum' }],
  target: { sheetId: 'sheet-1', startRow: 0, startColumn: 6 },
}

describe('PivotPanel command bindings', () => {
  it('routes every editing mutation through snapshot undo records', () => {
    const records: PivotUndoRecord[] = []
    const controller = new PivotCommandController(manager(), { push: (record) => records.push(record) })
    expect(controller.add(SPEC)).not.toBeNull()
    records.length = 0
    const commands = createPivotPanelCommandBindings(controller, SPEC.id)

    expect(commands.setRowFields(['Product'])).toBe(true)
    expect(commands.setColumnFields(['Quarter'])).toBe(true)
    expect(commands.setValueFields([{ field: 'Sales', agg: 'avg' }])).toBe(true)
    expect(commands.setFilters({ Region: ['West'] })).toBe(true)
    expect(commands.update({ grandTotals: false })).toBe(true)
    expect(commands.remove()).toBe(true)

    expect(records.map(({ label }) => label)).toEqual([
      'Set pivot row fields',
      'Set pivot column fields',
      'Set pivot value fields',
      'Set pivot filters',
      'Update pivot',
      'Remove pivot',
    ])
    expect(records[0].before.pivots[0].rows).toEqual(['Region'])
    expect(records.at(-1)?.before.pivots[0].grandTotals).toBe(false)
    expect(records.at(-1)?.after.pivots).toEqual([])
  })
})
