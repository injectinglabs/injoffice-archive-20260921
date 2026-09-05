import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { ConnectorManager } from './manager'
import { RangePreprocessPipeline } from './preprocess'
import type { ConnectorScheduler, ConnectorSpec, SourceFetcher } from './types'

function spec(patch: Partial<ConnectorSpec> = {}): ConnectorSpec {
  return {
    id: 'weather',
    name: 'Weather',
    source: { kind: 'http', url: 'https://data.example/weather.json', format: 'json' },
    target: { sheetId: 'sheet-1', startRow: 2, startColumn: 3 },
    refresh: 'manual',
    ...patch,
  }
}

function harness() {
  const writes: Array<{ row: number; column: number; rows: number; columns: number; values: unknown }> = []
  const sheet = {
    getRange: (row: number, column: number, rows: number, columns: number) => ({
      setValues: (values: unknown) => writes.push({ row, column, rows, columns, values }),
    }),
  }
  const workbook = { getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null }
  const api = { getActiveWorkbook: () => workbook } as unknown as FUniver
  return { api, writes }
}

describe('ConnectorManager lifecycle', () => {
  it('authorizes before fetching and exposes a structured denial', async () => {
    const { api, writes } = harness()
    const fetcher = vi.fn(async () => [[1]])
    const manager = new ConnectorManager(api, fetcher, { authorize: async () => false })
    manager.add(spec())
    expect(await manager.refresh('weather')).toBe('failed')
    expect(fetcher).not.toHaveBeenCalled()
    expect(writes).toHaveLength(0)
    expect(manager.status('weather')).toMatchObject({ refreshing: false, lastErrorCode: 'AUTHORIZATION_DENIED' })
  })

  it('uses and expires the memory cache while force bypasses it', async () => {
    const { api, writes } = harness()
    let now = 100
    const fetcher = vi.fn(async () => [[1, 'sun']])
    const manager = new ConnectorManager(api, fetcher, { now: () => now })
    manager.add(spec({ cache: { mode: 'memory', ttlMs: 50 } }))
    await manager.refresh('weather')
    await manager.refresh('weather')
    expect(fetcher).toHaveBeenCalledTimes(1)
    expect(manager.status('weather')?.cacheHit).toBe(true)
    await manager.refresh('weather', { force: true })
    expect(fetcher).toHaveBeenCalledTimes(2)
    now = 151
    await manager.refresh('weather')
    expect(fetcher).toHaveBeenCalledTimes(3)
    expect(writes).toHaveLength(4)
  })

  it('cancels an in-flight fetch and never writes a late result', async () => {
    const { api, writes } = harness()
    let release!: (value: unknown[][]) => void
    const fetcher: SourceFetcher = vi.fn(() => new Promise<unknown[][]>((resolve) => { release = resolve }))
    const manager = new ConnectorManager(api, fetcher)
    manager.add(spec())
    const pending = manager.refresh('weather')
    expect(manager.cancel('weather')).toBe(true)
    release([[99]])
    expect(await pending).toBe('canceled')
    expect(writes).toHaveLength(0)
    expect(manager.status('weather')?.lastErrorCode).toBe('ABORTED')
  })

  it('validates the declared schema before writing', async () => {
    const { api, writes } = harness()
    const manager = new ConnectorManager(api, async () => [['not-a-number']])
    manager.add(spec({ schema: { columns: [{ index: 0, type: 'number' }] } }))
    expect(await manager.refresh('weather')).toBe('failed')
    expect(manager.status('weather')?.lastErrorCode).toBe('INVALID_DATA')
    expect(writes).toHaveLength(0)
  })

  it('runs configured preprocessing with collaborative revision and fingerprint checks', async () => {
    const { api, writes } = harness()
    const preprocessing = new RangePreprocessPipeline()
    preprocessing.register({ id: 'trim', version: '1', order: 0, deterministic: true, process: (grid) => grid.map((row) => row.map((cell) => typeof cell === 'string' ? cell.trim() : cell)) })
    const manager = new ConnectorManager(api, async () => [[' Oslo ']], { preprocessing })
    manager.add(spec())
    expect(await manager.refresh('weather', { mode: 'collaborative' })).toBe('failed')
    expect(manager.status('weather')?.lastErrorCode).toBe('PREPROCESS_FAILED')
    expect(await manager.refresh('weather', { mode: 'collaborative', revision: 'rev-2', expectedPreprocessFingerprint: preprocessing.fingerprint() })).toBe('applied')
    expect(writes.at(-1)?.values).toEqual([[{ v: 'Oslo' }]])
  })

  it('schedules interval refresh and clears timers on removal', async () => {
    const { api } = harness()
    let callback!: () => void
    const scheduler: ConnectorScheduler = {
      setInterval: vi.fn((next) => { callback = next; return 'timer-1' }),
      clearInterval: vi.fn(),
    }
    const fetcher = vi.fn(async () => [[1]])
    const manager = new ConnectorManager(api, fetcher, { scheduler, now: () => 500 })
    manager.add(spec({ refresh: 'interval', schedule: { intervalMs: 1_000 } }))
    expect(manager.status('weather')?.nextRefreshTs).toBe(1_500)
    callback()
    await vi.waitFor(() => expect(fetcher).toHaveBeenCalledOnce())
    manager.remove('weather')
    expect(scheduler.clearInterval).toHaveBeenCalledWith('timer-1')
  })

  it('clears the previous extent before writing a smaller refresh', async () => {
    const { api, writes } = harness()
    const fetcher = vi.fn<SourceFetcher>()
      .mockResolvedValueOnce([[1, 2], [3, 4]])
      .mockResolvedValueOnce([[5]])
    const manager = new ConnectorManager(api, fetcher)
    manager.add(spec())
    await manager.refresh('weather')
    await manager.refresh('weather')
    expect(writes[1]).toMatchObject({ rows: 2, columns: 2, values: [[{ v: null }, { v: null }], [{ v: null }, { v: null }]] })
    expect(writes[2]).toMatchObject({ rows: 1, columns: 1, values: [[{ v: 5 }]] })
  })
})
