import type { FUniver } from '@univerjs/core/lib/facade'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ChartManager } from './manager'
import type { ChartSpec } from './types'
import { getChartOption } from './registry'

const A: ChartSpec = {
  id: 'ca',
  type: 'column',
  range: { sheetId: 's1', startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
}
const B: ChartSpec = {
  id: 'cb',
  type: 'line',
  range: { sheetId: 's1', startRow: 0, startColumn: 5, endRow: 3, endColumn: 7 },
}

function setRange(row: number, col: number, sheetId = 's1') {
  return {
    id: 'sheet.mutation.set-range-values',
    type: 2,
    params: { unitId: 'wb1', subUnitId: sheetId, cellValue: { [row]: { [col]: { v: 1 } } } },
  }
}

function fakeUniver() {
  const listeners: Array<(info: unknown, options?: unknown) => void> = []
  const reads: Array<{ startRow: number; startColumn: number }> = []
  const sheet = {
    getSheetId: () => 's1',
    getRange: (startRow: number, startColumn: number) => ({
      getValues: () => {
        reads.push({ startRow, startColumn })
        return [
          ['H', 'V'],
          ['a', 1],
        ]
      },
    }),
    addFloatDomToPosition: () => ({ id: 'dom', dispose: () => {} }),
  }
  const api = {
    onCommandExecuted: (cb: (info: unknown, options?: unknown) => void) => {
      listeners.push(cb)
      return { dispose: () => {} }
    },
    getActiveWorkbook: () => ({
      getSheetBySheetId: (id: string) => (id === 's1' ? sheet : null),
      getActiveSheet: () => sheet,
    }),
    fire(info: unknown, options?: unknown) {
      for (const l of listeners) l(info, options)
    },
  }
  return { api: api as unknown as FUniver, fire: api.fire.bind(api), reads }
}

describe('ChartManager range-aware invalidation', () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })
  afterEach(() => {
    vi.useRealTimers()
  })

  it('charts underlying numbers when the host also returns formatted display values', () => {
    const getValues = vi.fn(() => [['Company', 'USD billions'], ['Sample', '$5,331.0bn']])
    const getRawValues = vi.fn(() => [['Company', 'USD billions'], ['Sample', 5331]])
    const api = { getActiveWorkbook: () => ({ getSheetBySheetId: () => ({
      getRange: () => ({ getValues, getRawValues }),
      addFloatDomToPosition: () => ({ id: 'raw-chart', dispose: () => {} }),
    }) }) } as unknown as FUniver
    const manager = new ChartManager(api)
    expect(manager.add(A)).toBe(true)
    expect(getRawValues).toHaveBeenCalled()
    expect(getValues).not.toHaveBeenCalled()
    expect(getChartOption(A.id)?.series).toEqual(expect.arrayContaining([expect.objectContaining({ data: [5331] })]))
    manager.stop()
  })

  it('refreshes only the chart whose source intersects the edited cells', () => {
    const { api, fire, reads } = fakeUniver()
    const manager = new ChartManager(api)
    manager.start()
    expect(manager.add(A)).toBe(true)
    expect(manager.add(B)).toBe(true)
    const afterMount = reads.length

    fire(setRange(1, 0))
    vi.advanceTimersByTime(120)
    const afterA = reads.slice(afterMount)
    expect(afterA).toEqual([{ startRow: 0, startColumn: 0 }])

    fire(setRange(1, 6))
    vi.advanceTimersByTime(120)
    const afterB = reads.slice(afterMount + afterA.length)
    expect(afterB).toEqual([{ startRow: 0, startColumn: 5 }])
    manager.stop()
  })

  it('does not rebuild charts on selection commands', () => {
    const { api, fire, reads } = fakeUniver()
    const manager = new ChartManager(api)
    manager.start()
    manager.add(A)
    manager.add(B)
    const afterMount = reads.length
    fire({ id: 'sheet.operation.set-selections', params: { subUnitId: 's1' } })
    vi.advanceTimersByTime(120)
    expect(reads.length).toBe(afterMount)
    manager.stop()
  })

  it('coalesces a burst that hits two charts into one refresh each', () => {
    const { api, fire, reads } = fakeUniver()
    const manager = new ChartManager(api)
    manager.start()
    manager.add(A)
    manager.add(B)
    const afterMount = reads.length
    fire(setRange(1, 0))
    fire(setRange(2, 6))
    vi.advanceTimersByTime(120)
    const after = reads.slice(afterMount)
    expect(after).toEqual([
      { startRow: 0, startColumn: 0 },
      { startRow: 0, startColumn: 5 },
    ])
    manager.stop()
  })

  it('refreshes every chart on the sheet after a row insert', () => {
    const { api, fire, reads } = fakeUniver()
    const manager = new ChartManager(api)
    manager.start()
    manager.add(A)
    manager.add(B)
    const afterMount = reads.length
    fire({
      id: 'sheet.mutation.insert-row',
      type: 2,
      params: { subUnitId: 's1', range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 } },
    })
    vi.advanceTimersByTime(120)
    expect(reads.slice(afterMount)).toEqual([
      { startRow: 0, startColumn: 0 },
      { startRow: 0, startColumn: 5 },
    ])
    manager.stop()
  })
})

describe('ChartManager layering', () => {
  it('keeps a durable back-to-front order and asks the host before committing', () => {
    const { api } = fakeUniver()
    const applied: Array<{ sheetId: string; ids: readonly string[] }> = []
    const manager = new ChartManager(api, undefined, {
      layerHost: { setChartOrder: (sheetId, ids) => { applied.push({ sheetId, ids }); return true } },
    })
    const C: ChartSpec = { ...A, id: 'cc' }
    manager.add(A)
    manager.add(B)
    manager.add(C)

    expect(manager.layer('ca', 'bringToFront')).toBe(true)
    expect(manager.list().map(({ id }) => id)).toEqual(['cb', 'cc', 'ca'])
    expect(applied).toEqual([{ sheetId: 's1', ids: ['cb', 'cc', 'ca'] }])
    expect(manager.layer('ca', 'sendBackward')).toBe(true)
    expect(manager.list().map(({ id }) => id)).toEqual(['cb', 'ca', 'cc'])
    expect(manager.layer('cb', 'sendBackward')).toBe(false)
  })

  it('does not mutate order when the visual host refuses or throws', () => {
    const { api } = fakeUniver()
    let throwNow = false
    const manager = new ChartManager(api, undefined, {
      layerHost: { setChartOrder: () => { if (throwNow) throw new Error('not mounted'); return false } },
    })
    manager.add(A)
    manager.add(B)
    expect(manager.layer('ca', 'bringForward')).toBe(false)
    expect(manager.list().map(({ id }) => id)).toEqual(['ca', 'cb'])
    throwNow = true
    expect(manager.layer('ca', 'bringForward')).toBe(false)
    expect(manager.list().map(({ id }) => id)).toEqual(['ca', 'cb'])
  })
})
