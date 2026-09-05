import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it } from 'vitest'
import { PivotCommandController, type PivotUndoRecord } from './commands'
import { PivotManager } from './manager'
import type { PivotSpec } from './types'

const GRID = [
  ['Region', 'Product', 'Quarter', 'Sales'],
  ['West', 'Widget', 'Q1', 20],
  ['East', 'Gadget', 'Q2', 35],
]

function harness() {
  const writes: unknown[] = []
  const sheet = {
    getSheetId: () => 'sheet-1',
    getActiveRange: () => ({ getRange: () => ({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 3 }) }),
    getRange: () => ({ getValues: () => GRID, setValues: (value: unknown) => writes.push(value) }),
  }
  const workbook = {
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getActiveRange: () => null,
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }
  return { api: { getActiveWorkbook: () => workbook } as unknown as FUniver, writes }
}

function spec(overrides: Partial<PivotSpec> = {}): PivotSpec {
  return {
    id: 'pivot-1',
    source: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 3 },
    rows: ['Region'],
    columns: [],
    values: [{ field: 'Sales', agg: 'sum' }],
    target: { sheetId: 'sheet-1', startRow: 0, startColumn: 6 },
    ...overrides,
  }
}

describe('PivotCommandController', () => {
  it('records lifecycle and field operations as atomic complete snapshots', () => {
    const host = harness()
    const records: PivotUndoRecord[] = []
    const subject = new PivotCommandController(new PivotManager(host.api), { push: (record) => records.push(record) })
    const handle = subject.add(spec())
    expect(handle?.setRowFields(['Product'])).toBe(true)
    expect(handle?.setColumnFields(['Quarter'])).toBe(true)
    expect(handle?.setValueFields([{ field: 'Sales', agg: 'avg' }])).toBe(true)
    expect(handle?.setFilters({ Quarter: ['Q2'] })).toBe(true)
    expect(handle?.setFilters(undefined)).toBe(true)
    expect(handle?.setPageFields([{ field: 'Region', selectedItem: 'West' }])).toBe(true)
    expect(handle?.setMemberFilters([{ field: 'Product', mode: 'exclude', values: ['Gadget'] }])).toBe(true)
    expect(handle?.setSorts([{ field: 'Quarter', direction: 'descending' }])).toBe(true)
    expect(handle?.remove()).toBe(true)
    expect(records.map(({ label }) => label)).toEqual([
      'Add pivot', 'Set pivot row fields', 'Set pivot column fields', 'Set pivot value fields',
      'Set pivot filters', 'Set pivot filters',
      'Set pivot page fields', 'Set pivot member filters', 'Set pivot sorts', 'Remove pivot',
    ])
    expect(records[0].before.pivots).toEqual([])
    expect(records.at(-1)?.before.pivots[0]).toMatchObject({
      rows: ['Product'], columns: ['Quarter'], values: [{ field: 'Sales', agg: 'avg' }],
    })
    expect(records.at(-1)?.after.pivots).toEqual([])
    expect(records[5].before.pivots[0].filters).toEqual({ Quarter: ['Q2'] })
    expect(records[5].after.pivots[0].filters).toBeUndefined()
  })

  it('restores exact native identity and refuses identity replacement', () => {
    const host = harness()
    const subject = new PivotCommandController(new PivotManager(host.api))
    const nativeIdentity = { part: 'xl/pivotTables/pivotTable7.xml' }
    subject.add(spec({ nativeIdentity }))
    const snapshot = subject.snapshot()
    expect(subject.update('pivot-1', { nativeIdentity: { part: 'xl/pivotTables/pivotTable9.xml' }, grandTotals: false } as never)).toBe(true)
    expect(subject.getById('pivot-1')?.value).toMatchObject({ nativeIdentity, grandTotals: false })
    expect(subject.restore(snapshot)).toBe(true)
    expect(subject.getById('pivot-1')?.value).toMatchObject({ nativeIdentity })
    expect(subject.getById('pivot-1')?.value?.grandTotals).toBeUndefined()
  })

  it('creates from the active selection and rolls back invalid snapshots', () => {
    const host = harness()
    const subject = new PivotCommandController(new PivotManager(host.api))
    expect(subject.createFromSelection()?.value).toMatchObject({ rows: ['Region'], values: [{ field: 'Sales', agg: 'sum' }] })
    const before = subject.snapshot()
    expect(subject.restore({ version: 1, pivots: [spec({ source: { ...spec().source, sheetId: 'missing' } })] })).toBe(false)
    expect(subject.snapshot()).toEqual(before)
  })

  it('rejects malformed and host-incompatible field configurations', () => {
    const host = harness()
    const subject = new PivotCommandController(new PivotManager(host.api))
    expect(subject.add(spec({ id: '', nativeIdentity: { part: '../pivot.xml' } }))).toBeNull()
    expect(subject.add(spec())).not.toBeNull()
    expect(subject.setColumnFields('pivot-1', ['Quarter', 'Product'])).toBe(false)
    expect(subject.setColumnFields('pivot-1', ['Region'])).toBe(false)
    expect(subject.setValueFields('pivot-1', [{ field: 'Missing', agg: 'sum' }])).toBe(false)
    expect(subject.setPageFields('pivot-1', [{ field: 'Region' }])).toBe(false)
    expect(subject.setSorts('pivot-1', [{ field: 'Sales', direction: 'ascending' }])).toBe(false)
    expect(subject.update('pivot-1', { mystery: true } as never)).toBe(false)
    expect(subject.restore({ version: 1, pivots: [null] } as never)).toBe(false)
  })
})
