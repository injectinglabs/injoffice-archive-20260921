import type { ICellData } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { ConnectorCommandController, isCredentialFreeConnectorSpec, type ConnectorUndoRecord } from './commands'
import { ConnectorManager } from './manager'
import type { ConnectorSpec, SourceFetcher } from './types'

function spec(patch: Partial<ConnectorSpec> = {}): ConnectorSpec {
  return {
    id: 'weather',
    name: 'Weather',
    source: { kind: 'http', url: '/gateway/weather', format: 'json' },
    target: { sheetId: 'sheet-1', startRow: 0, startColumn: 0 },
    refresh: 'manual',
    ...patch,
  }
}

function harness(fetcher: SourceFetcher) {
  const cells: ICellData[][] = Array.from({ length: 8 }, () =>
    Array.from({ length: 8 }, () => ({ v: null })))
  let failNextDataWrite = false
  const sheet = {
    getSheetId: () => 'sheet-1',
    getActiveRange: () => ({ getRange: () => ({ startRow: 0, startColumn: 0 }) }),
    getRange: (startRow: number, startColumn: number, rowCount: number, columnCount: number) => ({
      getValues: () => Array.from({ length: rowCount }, (_, row) =>
        Array.from({ length: columnCount }, (_, column) => structuredClone(cells[startRow + row][startColumn + column].v ?? null))),
      getFormulas: () => Array.from({ length: rowCount }, (_, row) =>
        Array.from({ length: columnCount }, (_, column) => cells[startRow + row][startColumn + column].f ?? null)),
      setValues: (values: ICellData[][]) => {
        const hasData = values.some((row) => row.some((cell) => cell.v !== null))
        if (failNextDataWrite && hasData) {
          failNextDataWrite = false
          throw new Error('simulated worksheet write failure')
        }
        for (let row = 0; row < rowCount; row++) {
          for (let column = 0; column < columnCount; column++) {
            cells[startRow + row][startColumn + column] = structuredClone(values[row][column])
          }
        }
      },
    }),
  }
  const workbook = {
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getActiveRange: () => sheet.getActiveRange(),
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }
  const api = { getActiveWorkbook: () => workbook } as unknown as FUniver
  const undo: ConnectorUndoRecord[] = []
  const manager = new ConnectorManager(api, fetcher, { now: () => 1_000 })
  const controller = new ConnectorCommandController(manager, { push: (record) => undo.push(record) })
  return {
    cells,
    manager,
    controller,
    undo,
    failNextDataWrite: () => { failNextDataWrite = true },
  }
}

describe('ConnectorCommandController', () => {
  it('creates and refreshes as one undoable operation without refetching on redo', async () => {
    const fetcher = vi.fn(async () => [['fresh', 2], [3, 4]])
    const subject = harness(fetcher)
    subject.cells[0][0] = { v: 'original' }
    const handle = await subject.controller.createAtSelection({
      name: 'Weather', source: { kind: 'http', url: '/gateway/weather', format: 'json' }, refresh: 'manual',
    })
    expect(handle?.value?.target).toEqual({ sheetId: 'sheet-1', startRow: 0, startColumn: 0 })
    expect(subject.cells[0][0].v).toBe('fresh')
    expect(subject.undo).toHaveLength(1)

    expect(subject.controller.restore(subject.undo[0].before)).toBe(true)
    expect(subject.manager.list()).toEqual([])
    expect(subject.cells[0][0].v).toBe('original')
    expect(subject.controller.restore(subject.undo[0].after)).toBe(true)
    expect(subject.manager.list()).toHaveLength(1)
    expect(subject.cells[1][1].v).toBe(4)
    expect(fetcher).toHaveBeenCalledOnce()
  })

  it('captures growth beyond the previous extent and replays refresh cells without network I/O', async () => {
    const fetcher = vi.fn<SourceFetcher>()
      .mockResolvedValueOnce([[1]])
      .mockResolvedValueOnce([[2, 3], [4, 5]])
    const subject = harness(fetcher)
    subject.cells[0][1] = { v: 2, f: '=1+1' }
    subject.cells[1][0] = { v: 'below' }
    subject.manager.add(spec())
    await subject.manager.refresh('weather')
    expect(await subject.controller.refresh('weather')).toBe('applied')
    const record = subject.undo[0]
    expect(record.before.ranges[0]).toMatchObject({ rowCount: 2, columnCount: 2 })
    expect(subject.controller.restore(record.before)).toBe(true)
    expect(subject.cells[0][0].v).toBe(1)
    expect(subject.cells[0][1]).toEqual({ f: '=1+1' })
    expect(subject.cells[1][0].v).toBe('below')
    expect(subject.controller.restore(record.after)).toBe(true)
    expect(subject.cells[1][1].v).toBe(5)
    expect(fetcher).toHaveBeenCalledTimes(2)
  })

  it('rolls back a partial shrink write and does not add undo when refresh fails', async () => {
    const fetcher = vi.fn<SourceFetcher>()
      .mockResolvedValueOnce([[1, 2], [3, 4]])
      .mockResolvedValueOnce([[9]])
    const subject = harness(fetcher)
    subject.manager.add(spec())
    await subject.manager.refresh('weather')
    subject.failNextDataWrite()
    expect(await subject.controller.refresh('weather')).toBe('failed')
    expect(subject.cells.slice(0, 2).map((row) => row.slice(0, 2).map((cell) => cell.v)))
      .toEqual([[1, 2], [3, 4]])
    expect(subject.manager.status('weather')).toMatchObject({ rowCount: 2, columnCount: 2 })
    expect(subject.undo).toEqual([])
  })

  it('undoes target updates and removal together with their cleared cells', async () => {
    const subject = harness(async () => [[7, 8]])
    const added = subject.controller.add(spec())
    expect(added).not.toBeNull()
    await subject.manager.refresh('weather')
    subject.undo.length = 0

    expect(subject.controller.update('weather', { target: { sheetId: 'sheet-1', startRow: 3, startColumn: 3 } })).toBe(true)
    expect(subject.cells[0][0].v).toBeNull()
    expect(subject.controller.restore(subject.undo[0].before)).toBe(true)
    expect(subject.cells[0][0].v).toBe(7)
    expect(subject.manager.get('weather')?.target.startRow).toBe(0)
    expect(subject.controller.restore(subject.undo[0].after)).toBe(true)
    expect(subject.cells[0][0].v).toBeNull()

    subject.undo.length = 0
    await subject.controller.refresh('weather')
    subject.undo.length = 0
    expect(subject.controller.remove('weather')).toBe(true)
    expect(subject.manager.get('weather')).toBeUndefined()
    expect(subject.cells[3][3].v).toBeNull()
    expect(subject.controller.restore(subject.undo[0].before)).toBe(true)
    expect(subject.manager.get('weather')).toBeDefined()
    expect(subject.cells[3][3].v).toBe(7)
  })

  it('refuses secret-bearing specs, hostile policy failures, and malformed snapshots without mutation', () => {
    const subject = harness(async () => [[1]])
    expect(isCredentialFreeConnectorSpec(spec())).toBe(true)
    expect(isCredentialFreeConnectorSpec(spec({ source: { kind: 'http', url: 'https://user:pass@example.test/data', format: 'json' } }))).toBe(false)
    expect(isCredentialFreeConnectorSpec(spec({ source: { kind: 'http', url: '/gateway?token=secret', format: 'json' } }))).toBe(false)
    expect(isCredentialFreeConnectorSpec(spec({ source: { kind: 'http', url: '/gateway', format: 'json', authorization: 'Bearer secret' } as never }))).toBe(false)
    expect(subject.controller.add(spec({ source: { kind: 'http', url: '/gateway?key=secret', format: 'json' } }))).toBeNull()
    expect(subject.controller.restore({ version: 1, manager: { version: 1, connectors: [] }, ranges: [{
      sheetId: 'sheet-1', startRow: 0, startColumn: 0, rowCount: 2, columnCount: 2, values: [[{ v: 1 }]],
    }] })).toBe(false)
    expect(subject.manager.list()).toEqual([])

    const denied = new ConnectorCommandController(subject.manager, undefined, { canPersist: () => { throw new Error('policy unavailable') } })
    expect(() => denied.add(spec())).not.toThrow()
    expect(denied.add(spec())).toBeNull()
  })

  it('rolls back the currently attempted range when a host write fails partway', () => {
    const cells: ICellData[][] = [[{ v: 'left' }, { v: 'right' }]]
    let failPartway = true
    const sheet = {
      getSheetId: () => 'sheet-1',
      getRange: () => ({
        getValues: () => [[cells[0][0].v, cells[0][1].v]],
        getFormulas: () => [[null, null]],
        setValues: (values: ICellData[][]) => {
          cells[0][0] = structuredClone(values[0][0])
          if (failPartway) { failPartway = false; throw new Error('partial host write') }
          cells[0][1] = structuredClone(values[0][1])
        },
      }),
    }
    const api = { getActiveWorkbook: () => ({ getActiveSheet: () => sheet, getSheetBySheetId: () => sheet }) } as unknown as FUniver
    const controller = new ConnectorCommandController(new ConnectorManager(api, async () => []))
    expect(controller.restore({
      version: 1,
      manager: { version: 1, connectors: [] },
      ranges: [{ sheetId: 'sheet-1', startRow: 0, startColumn: 0, rowCount: 1, columnCount: 2, values: [[{ v: 'new-left' }, { v: 'new-right' }]] }],
    })).toBe(false)
    expect(cells).toEqual([[{ v: 'left' }, { v: 'right' }]])
  })
})
