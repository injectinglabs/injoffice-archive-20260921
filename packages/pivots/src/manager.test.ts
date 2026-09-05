import type { FUniver } from '@univerjs/core/lib/facade'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { PivotManager } from './manager'
import type { PivotSpec } from './types'

const GRID = [
  ['Region', 'Sales'],
  ['West', 20],
  ['East', 35],
]

const P1: PivotSpec = {
  id: 'p1',
  source: { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
  rows: ['Region'],
  columns: [],
  values: [{ field: 'Sales', agg: 'sum' }],
  target: { sheetId: 's1', startRow: 0, startColumn: 5 },
}

const P2: PivotSpec = {
  id: 'p2',
  source: { sheetId: 's1', startRow: 10, startColumn: 0, endRow: 12, endColumn: 1 },
  rows: ['Region'],
  columns: [],
  values: [{ field: 'Sales', agg: 'sum' }],
  target: { sheetId: 's1', startRow: 10, startColumn: 5 },
}

const P3: PivotSpec = {
  id: 'p3',
  source: { sheetId: 's2', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
  rows: ['Region'],
  columns: [],
  values: [{ field: 'Sales', agg: 'sum' }],
  target: { sheetId: 's2', startRow: 20, startColumn: 5 },
}

/** P2 sourced from P1's baked target (pivot-of-pivot). */
const NESTED: PivotSpec = {
  id: 'nested',
  source: { sheetId: 's1', startRow: 0, startColumn: 5, endRow: 3, endColumn: 6 },
  rows: ['Region'],
  columns: [],
  values: [{ field: 'Sales', agg: 'sum' }],
  target: { sheetId: 's1', startRow: 0, startColumn: 12 },
}

function setRange(row: number, col: number, sheetId = 's1') {
  return {
    id: 'sheet.mutation.set-range-values',
    type: 2,
    params: { unitId: 'wb1', subUnitId: sheetId, cellValue: { [row]: { [col]: { v: 99 } } } },
  }
}

function fakeUniver() {
  const listeners: Array<(info: unknown, options?: unknown) => void> = []
  const bakes: Array<{ startRow: number; startColumn: number }> = []
  const sheet = {
    getSheetId: () => 's1',
    getRange: (startRow: number, startColumn: number) => ({
      getValues: () => GRID,
      setValues: () => {
        bakes.push({ startRow, startColumn })
      },
    }),
  }
  const api = {
    onCommandExecuted: (cb: (info: unknown, options?: unknown) => void) => {
      listeners.push(cb)
      return { dispose: () => {} }
    },
    getActiveWorkbook: () => ({
      getSheetBySheetId: (id: string) => (id === 's1' || id === 's2' ? sheet : null),
      getActiveSheet: () => sheet,
    }),
    fire(info: unknown, options?: unknown) {
      for (const l of listeners) l(info, options)
    },
  }
  return { api: api as unknown as FUniver, fire: api.fire.bind(api), bakes }
}

describe('PivotManager range-aware invalidation', () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })
  afterEach(() => {
    vi.useRealTimers()
  })

  it('builds lifecycle requests from the managed hydrated identity', () => {
    const { api } = fakeUniver()
    const manager = new PivotManager(api)
    manager.add({ ...P1, nativeIdentity: { part: 'xl/pivotTables/pivotTable4.xml' } })
    expect(manager.nativeRemoveRequest('p1').request).toEqual({
      operation: 'remove',
      identity: { part: 'xl/pivotTables/pivotTable4.xml' },
    })
    expect(manager.nativeUpdateRequest('p1', {
      sheetNameOf: (id) => id === 's1' ? 'Data' : null,
      headerRowOf: () => GRID[0],
    }).request).toMatchObject({
      operation: 'update',
      identity: { part: 'xl/pivotTables/pivotTable4.xml' },
      pivot: { sourceSheetName: 'Data', sourceRef: 'A1:B3' },
    })
    expect(manager.nativeRemoveRequest('missing').skipped[0]).toContain('not managed')
  })

  it('rebakes only the pivot whose source intersects the edit', () => {
    const { api, fire, bakes } = fakeUniver()
    const manager = new PivotManager(api)
    manager.start()
    expect(manager.add(P1)).toBe(true)
    expect(manager.add(P2)).toBe(true)
    const afterMount = bakes.length
    expect(afterMount).toBe(2)

    fire(setRange(1, 1))
    vi.advanceTimersByTime(200)
    expect(bakes.slice(afterMount)).toEqual([{ startRow: 0, startColumn: 5 }])

    fire(setRange(11, 0))
    vi.advanceTimersByTime(200)
    expect(bakes.slice(afterMount + 1)).toEqual([{ startRow: 10, startColumn: 5 }])
    manager.stop()
  })

  it('does not rebake when only the baked target is edited', () => {
    const { api, fire, bakes } = fakeUniver()
    const manager = new PivotManager(api)
    manager.start()
    manager.add(P1)
    const afterMount = bakes.length
    fire(setRange(0, 6))
    vi.advanceTimersByTime(200)
    expect(bakes.length).toBe(afterMount)
    manager.stop()
  })

  it('skips fromCollab payloads to avoid rebake ping-pong', () => {
    const { api, fire, bakes } = fakeUniver()
    const manager = new PivotManager(api)
    manager.start()
    manager.add(P1)
    const afterMount = bakes.length
    fire(setRange(1, 1), { fromCollab: true })
    vi.advanceTimersByTime(200)
    expect(bakes.length).toBe(afterMount)
    manager.stop()
  })

  it('rebakes a nested pivot whose source is another pivot\'s target', () => {
    const { api, fire, bakes } = fakeUniver()
    const manager = new PivotManager(api)
    manager.start()
    expect(manager.add(P1)).toBe(true)
    expect(manager.add(NESTED)).toBe(true)
    const afterMount = bakes.length
    fire(setRange(1, 1))
    vi.advanceTimersByTime(200)
    expect(bakes.slice(afterMount)).toEqual([
      { startRow: 0, startColumn: 5 },
      { startRow: 0, startColumn: 12 },
    ])
    manager.stop()
  })

  it('does not rebake on selection commands', () => {
    const { api, fire, bakes } = fakeUniver()
    const manager = new PivotManager(api)
    manager.start()
    manager.add(P1)
    manager.add(P2)
    const afterMount = bakes.length
    fire({ id: 'sheet.operation.set-selections', params: { subUnitId: 's1' } })
    vi.advanceTimersByTime(200)
    expect(bakes.length).toBe(afterMount)
    manager.stop()
  })

  it('coalesces a burst that hits two sources into one rebake each', () => {
    const { api, fire, bakes } = fakeUniver()
    const manager = new PivotManager(api)
    manager.start()
    manager.add(P1)
    manager.add(P2)
    const afterMount = bakes.length
    fire(setRange(1, 1))
    fire(setRange(11, 0))
    vi.advanceTimersByTime(200)
    expect(bakes.slice(afterMount)).toEqual([
      { startRow: 0, startColumn: 5 },
      { startRow: 10, startColumn: 5 },
    ])
    manager.stop()
  })

  it('rebakes every pivot on the edited sheet after a row insert, not another sheet', () => {
    const { api, fire, bakes } = fakeUniver()
    const manager = new PivotManager(api)
    manager.start()
    manager.add(P1)
    manager.add(P2)
    manager.add(P3)
    const afterMount = bakes.length
    fire({
      id: 'sheet.mutation.insert-row',
      type: 2,
      params: { subUnitId: 's1', range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 } },
    })
    vi.advanceTimersByTime(200)
    expect(bakes.slice(afterMount)).toEqual([
      { startRow: 0, startColumn: 5 },
      { startRow: 10, startColumn: 5 },
    ])
    manager.stop()
  })
})
