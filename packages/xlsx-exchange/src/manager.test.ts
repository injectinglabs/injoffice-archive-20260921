import { describe, expect, it, vi } from 'vitest'
import { XlsxExchangeManager } from './manager'
import {
  XlsxExchangeError,
  type XlsxExchangeCodec,
  type XlsxExchangeContext,
  type XlsxExchangeStorage,
  type XlsxExchangeWorkspace,
} from './types'

interface Snapshot { readonly cells: readonly string[] }

function createHarness(codec: XlsxExchangeCodec<Snapshot> = {}) {
  const storage = {
    read: vi.fn(async (_source: string, _context: XlsxExchangeContext) => ({ bytes: new Uint8Array([1, 2, 3]), name: 'stored.xlsx' })),
    saveAs: vi.fn(async (_file, _destination: string | undefined, _context: XlsxExchangeContext) => ({ location: 'download://saved.xlsx' })),
  } satisfies XlsxExchangeStorage<string, string>
  const workspace = {
    loadSnapshot: vi.fn(async () => ({ unitId: 'loaded-snapshot' })),
    loadServerUnit: vi.fn(async (unitId: string) => ({ unitId })),
    captureSnapshot: vi.fn(async () => ({ cells: ['captured'] })),
  } satisfies XlsxExchangeWorkspace<Snapshot>
  return {
    storage,
    workspace,
    manager: new XlsxExchangeManager<Snapshot, string, string>({ codec, storage, workspace }),
  }
}

describe('XlsxExchangeManager', () => {
  it('imports a byte snapshot before loading it and isolates caller bytes', async () => {
    const order: string[] = []
    let release!: () => void
    const wait = new Promise<void>((resolve) => { release = resolve })
    const codec: XlsxExchangeCodec<Snapshot> = {
      importSnapshot: vi.fn(async (file) => {
        order.push(`convert:${file.bytes.join(',')}`)
        await wait
        return { snapshot: { cells: ['A1'] }, revision: 'sha256:one' }
      }),
    }
    const { manager, workspace } = createHarness(codec)
    workspace.loadSnapshot.mockImplementation(async () => {
      order.push('load')
      return { unitId: 'book-1' }
    })
    const bytes = new Uint8Array([1, 2, 3])
    const job = manager.importSnapshot({ source: { kind: 'bytes', bytes, name: 'input.xlsx' }, load: { mode: 'new-unit' } })
    bytes.fill(9)
    await Promise.resolve()
    expect(workspace.loadSnapshot).not.toHaveBeenCalled()
    release()

    await expect(job.result).resolves.toMatchObject({ loadedUnitId: 'book-1', sourceName: 'input.xlsx' })
    expect(order).toEqual(['convert:1,2,3', 'load'])
  })

  it('does not touch the workbook when conversion fails', async () => {
    const { manager, workspace } = createHarness({
      importSnapshot: async () => { throw new Error('unsupported chart part') },
    })
    const job = manager.importSnapshot({
      source: { kind: 'bytes', bytes: new Uint8Array([1]) },
      load: { mode: 'replace', targetUnitId: 'book-1' },
    })
    await expect(job.result).rejects.toMatchObject({ code: 'IMPORT_FAILED', details: { phase: 'converting', retryable: true } })
    expect(workspace.loadSnapshot).not.toHaveBeenCalled()
    expect(job.snapshot().state).toBe('failed')
  })

  it('requires a target ID for safe replacement', async () => {
    const { manager, workspace } = createHarness({ importSnapshot: async () => ({ snapshot: { cells: [] } }) })
    const job = manager.importSnapshot({
      source: { kind: 'bytes', bytes: new Uint8Array([1]) },
      load: { mode: 'replace' },
    })
    await expect(job.result).rejects.toMatchObject({ code: 'INVALID_REQUEST' })
    expect(workspace.loadSnapshot).not.toHaveBeenCalled()
  })

  it('cancels without crossing the load boundary', async () => {
    let context!: XlsxExchangeContext
    let finish!: () => void
    let entered!: () => void
    const started = new Promise<void>((resolve) => { entered = resolve })
    const codec: XlsxExchangeCodec<Snapshot> = {
      importSnapshot: (_file, ctx) => new Promise((resolve) => {
        context = ctx
        finish = () => resolve({ snapshot: { cells: [] } })
        entered()
      }),
    }
    const { manager, workspace } = createHarness(codec)
    const job = manager.importSnapshot({ source: { kind: 'bytes', bytes: new Uint8Array([1]) }, load: { mode: 'new-unit' } })
    await started
    job.cancel('tab closed')
    expect(context.signal.aborted).toBe(true)
    finish()
    await expect(job.result).rejects.toMatchObject({ code: 'CANCELED' })
    expect(job.snapshot().state).toBe('canceled')
    expect(workspace.loadSnapshot).not.toHaveBeenCalled()
  })

  it('imports and optionally loads a server unit', async () => {
    const { manager, workspace } = createHarness({
      importServerUnit: async (_file, context) => {
        context.report({ phase: 'transferring', fraction: 0.7, loadedBytes: 3, totalBytes: 3 })
        return { unitId: 'server-42', revision: 'rev:one' }
      },
    })
    const job = manager.importServerUnit({
      source: { kind: 'storage', ref: 'inbox/item' },
      load: { mode: 'new-unit' },
    })
    await expect(job.result).resolves.toMatchObject({ unitId: 'server-42', loadedUnitId: 'server-42', sourceName: 'stored.xlsx' })
    expect(workspace.loadServerUnit).toHaveBeenCalledWith('server-42', { mode: 'new-unit' }, expect.objectContaining({ jobId: job.id }))
  })

  it('captures, converts, then saves a snapshot', async () => {
    const order: string[] = []
    const { manager, workspace, storage } = createHarness({
      exportSnapshot: async (snapshot) => {
        order.push(`convert:${snapshot.cells[0]}`)
        return { bytes: new Uint8Array([7, 8]), name: 'report.xlsx' }
      },
    })
    workspace.captureSnapshot.mockImplementation(async () => {
      order.push('capture')
      return { cells: ['captured'] }
    })
    storage.saveAs.mockImplementation(async (file) => {
      order.push(`save:${file.bytes.join(',')}`)
      file.bytes.fill(0)
      return { location: 'download://report.xlsx' }
    })
    const job = manager.exportSnapshot({ unitId: 'book-1', destination: 'downloads' })
    const result = await job.result
    expect(order).toEqual(['capture', 'convert:captured', 'save:7,8'])
    expect(result.file.bytes).toEqual(new Uint8Array([7, 8]))
    expect(storage.saveAs).toHaveBeenCalledWith(expect.objectContaining({ mediaType: expect.stringContaining('spreadsheetml') }), 'downloads', expect.anything())
  })

  it('does not invoke save-as when export fails', async () => {
    const { manager, storage } = createHarness({ exportSnapshot: async () => { throw new Error('native refusal') } })
    const job = manager.exportSnapshot({ snapshot: { cells: [] } })
    await expect(job.result).rejects.toMatchObject({ code: 'EXPORT_FAILED' })
    expect(storage.saveAs).not.toHaveBeenCalled()
  })

  it('bounds codec output before bytes cross the save-as boundary', async () => {
    const storage = {
      read: vi.fn(),
      saveAs: vi.fn(async () => ({ location: 'download://never.xlsx' })),
    } satisfies XlsxExchangeStorage<string, string>
    const manager = new XlsxExchangeManager<Snapshot, string, string>({
      codec: { exportServerUnit: async () => ({ bytes: new Uint8Array([1, 2, 3]) }) },
      storage,
      maxOutputBytes: 2,
    })
    await expect(manager.exportServerUnit({ unitId: 'unit-1' }).result).rejects.toMatchObject({ code: 'EXPORT_FAILED' })
    expect(storage.saveAs).not.toHaveBeenCalled()
    expect(() => new XlsxExchangeManager({ codec: {}, maxOutputBytes: 0 })).toThrow(TypeError)
  })

  it('validates XLSX identity and the expected immutable revision before save-as', async () => {
    const wrongRevision = createHarness({
      exportServerUnit: async () => ({ bytes: new Uint8Array([1]), revision: 'revision-2' }),
    })
    await expect(wrongRevision.manager.exportServerUnit({ unitId: 'unit-1', expectedRevision: 'revision-1' }).result)
      .rejects.toMatchObject({ code: 'EXPORT_FAILED' })
    expect(wrongRevision.storage.saveAs).not.toHaveBeenCalled()

    const wrongMediaType = createHarness({
      exportServerUnit: async () => ({ bytes: new Uint8Array([1]), mediaType: 'application/octet-stream' }),
    })
    await expect(wrongMediaType.manager.exportServerUnit({ unitId: 'unit-1' }).result)
      .rejects.toMatchObject({ code: 'EXPORT_FAILED' })
    expect(wrongMediaType.storage.saveAs).not.toHaveBeenCalled()
  })

  it('exports a server unit and reports immutable monotonic job progress', async () => {
    const fractions: number[] = []
    const { manager } = createHarness({
      exportServerUnit: async (_unitId, context) => {
        context.report({ phase: 'transferring', fraction: 0.2 }) // ignored regression
        context.report({ phase: 'converting', fraction: 0.75 })
        return { bytes: new Uint8Array([4]), revision: 'rev:two' }
      },
    })
    manager.subscribe((event) => fractions.push(event.job.progress.fraction))
    const job = manager.exportServerUnit({ unitId: 'unit-9' })
    await expect(job.result).resolves.toMatchObject({ file: { revision: 'rev:two' } })
    expect(fractions).toEqual([...fractions].sort((a, b) => a - b))
    expect(Object.isFrozen(job.snapshot())).toBe(true)
    expect(job.snapshot()).toMatchObject({ kind: 'export-server-unit', state: 'succeeded', progress: { phase: 'complete', fraction: 1 } })
  })

  it('fails closed for unconfigured directions', async () => {
    const { manager } = createHarness()
    const job = manager.exportServerUnit({ unitId: 'unit-1' })
    await expect(job.result).rejects.toBeInstanceOf(XlsxExchangeError)
    await expect(job.result).rejects.toMatchObject({ code: 'UNSUPPORTED_OPERATION', details: { retryable: false } })
  })
})
