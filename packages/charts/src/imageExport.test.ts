import { afterEach, describe, expect, it, vi } from 'vitest'
import { ChartImageExportError, ChartImageExportManager, type ChartImageExportEvent, type ChartImageRenderRequest } from './imageExport'
import type { ChartSpec } from './types'

const chart: ChartSpec = {
  id: 'chart-1',
  type: 'Column',
  title: 'Revenue',
  range: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
}

describe('ChartImageExportManager', () => {
  afterEach(() => vi.useRealTimers())

  it('runs a validated host export and emits a metadata-only lifecycle', async () => {
    const events: ChartImageExportEvent[] = []
    const render = vi.fn(async (_request: ChartImageRenderRequest) => ({ mediaType: 'image/png' as const, bytes: new Uint8Array([1, 2, 3]), width: 640, height: 360 }))
    const manager = new ChartImageExportManager((id) => id === chart.id ? chart : undefined, { render }, { createJobId: () => 'job-1' })
    manager.onEvent((event) => events.push(event))
    manager.onEvent(() => { throw new Error('observer failure') })

    const job = manager.start('chart-1', { pixelRatio: 2, backgroundColor: '#fff' })
    const result = await job.result
    expect(result).toEqual({ jobId: 'job-1', chartId: 'chart-1', format: 'png', mediaType: 'image/png', bytes: new Uint8Array([1, 2, 3]), width: 640, height: 360 })
    expect(render).toHaveBeenCalledWith(expect.objectContaining({ jobId: 'job-1', chart: expect.objectContaining({ id: 'chart-1' }), format: 'png', pixelRatio: 2, backgroundColor: '#fff' }), { signal: expect.any(AbortSignal) })
    expect(Object.isFrozen(render.mock.calls[0][0].chart.range)).toBe(true)
    expect(events).toEqual([
      { type: 'started', jobId: 'job-1', chartId: 'chart-1', format: 'png' },
      { type: 'completed', jobId: 'job-1', chartId: 'chart-1', format: 'png', mediaType: 'image/png', byteLength: 3, width: 640, height: 360 },
    ])
    expect(job.cancel()).toBe(false)
  })

  it('cancels even when a host ignores AbortSignal', async () => {
    const events: ChartImageExportEvent[] = []
    const manager = new ChartImageExportManager(() => chart, { render: () => new Promise(() => {}) }, { createJobId: () => 'job-cancel' })
    manager.onEvent((event) => events.push(event))
    const job = manager.start('chart-1', { format: 'svg' })
    const rejection = expect(job.result).rejects.toMatchObject({ code: 'CANCELED', message: 'no longer needed' })
    expect(job.cancel('no longer needed')).toBe(true)
    await rejection
    expect(events.map(({ type }) => type)).toEqual(['started', 'canceled'])
  })

  it('times out an uncooperative host and reports the distinct terminal state', async () => {
    vi.useFakeTimers()
    const events: ChartImageExportEvent[] = []
    const manager = new ChartImageExportManager(() => chart, { render: () => new Promise(() => {}) }, { timeoutMs: 25, createJobId: () => 'job-timeout' })
    manager.onEvent((event) => events.push(event))
    const job = manager.start('chart-1')
    const rejection = expect(job.result).rejects.toMatchObject({ code: 'TIMED_OUT' })
    await vi.advanceTimersByTimeAsync(25)
    await rejection
    expect(events.map(({ type }) => type)).toEqual(['started', 'timed-out'])
  })

  it('rejects malformed requests and host artifacts without leaking an active id', async () => {
    const manager = new ChartImageExportManager(
      (id) => id === chart.id ? chart : undefined,
      { render: async () => ({ mediaType: 'image/jpeg', bytes: new Uint8Array(), width: 0, height: 1 }) },
      { createJobId: () => 'repeatable', maxBytes: 4 },
    )
    expect(() => manager.start('missing')).toThrowError(ChartImageExportError)
    expect(() => manager.start('chart-1', { pixelRatio: 99 })).toThrowError(/pixelRatio/)
    expect(() => manager.start('chart-1', null as never)).toThrowError(/options must be an object/)
    expect(() => manager.start('chart-1', {}, {} as never)).toThrowError(/signal must implement AbortSignal/)
    await expect(manager.export('chart-1')).rejects.toMatchObject({ code: 'INVALID_RESULT' })
    await expect(manager.export('chart-1')).rejects.toMatchObject({ code: 'INVALID_RESULT' })
  })

  it('accepts caller cancellation and clones returned bytes', async () => {
    const source = new Uint8Array([7, 8])
    const manager = new ChartImageExportManager(() => chart, { render: async () => ({ mediaType: 'image/jpeg', bytes: source, width: 10, height: 10 }) })
    const result = await manager.export('chart-1', { format: 'jpeg' })
    source[0] = 0
    expect(result.bytes).toEqual(new Uint8Array([7, 8]))

    const abort = new AbortController()
    const blocked = new ChartImageExportManager(() => chart, { render: () => new Promise(() => {}) })
    const pending = blocked.export('chart-1', {}, abort.signal)
    const rejection = expect(pending).rejects.toMatchObject({ code: 'CANCELED' })
    abort.abort()
    await rejection
  })
})
