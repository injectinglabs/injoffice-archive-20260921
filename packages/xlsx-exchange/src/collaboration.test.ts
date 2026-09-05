import { describe, expect, it, vi } from 'vitest'
import {
  XLSX_COLLABORATIVE_IMPORT_PROTOCOL,
  XlsxCollaborativeImportCoordinator,
  XlsxCollaborativeImportError,
  XlsxCollaborativeImportRecoveryError,
  type XlsxCollaborationRoomState,
  type XlsxCollaborativeImportBoundary,
  type XlsxCollaborativeImportLease,
  type XlsxCollaborativeImportReceipt,
} from './collaboration'
import { XlsxExchangeManager } from './manager'
import type { XlsxExchangeCodec } from './types'

interface Snapshot { readonly cells: readonly string[] }

const source = { kind: 'bytes' as const, bytes: new Uint8Array([1, 2, 3]), name: 'team.xlsx' }

function room(overrides: Partial<XlsxCollaborationRoomState> = {}): XlsxCollaborationRoomState {
  return {
    room: 'room-1', unitId: 'old-unit', artifactVersion: 'version-7', logEpoch: 'epoch-7',
    headSequence: 12, savedSequence: 12, pendingOutbound: 0, phase: 'quiesced', ...overrides,
  }
}

function receipt(overrides: Partial<XlsxCollaborativeImportReceipt> = {}): XlsxCollaborativeImportReceipt {
  return {
    room: 'room-1', unitId: 'imported-unit', importedUnitId: 'imported-unit', artifactVersion: 'version-8',
    logEpoch: 'epoch-8', headSequence: 0, savedSequence: 0, pendingOutbound: 0, phase: 'active',
    transactionId: 'transaction-1', ...overrides,
  }
}

function harness(options: {
  codec?: XlsxExchangeCodec<Snapshot>
  boundary?: Partial<XlsxCollaborativeImportBoundary>
} = {}) {
  const calls: string[] = []
  const recovery = vi.fn()
  const codec = options.codec ?? {
    importServerUnit: vi.fn(async () => { calls.push('convert'); return { unitId: 'imported-unit', revision: 'native-3' } }),
  }
  const exchange = new XlsxExchangeManager<Snapshot>({ codec })
  const boundary: XlsxCollaborativeImportBoundary = {
    authorizeAndBegin: vi.fn(async ({ protocol }) => {
      calls.push(`authorize:${protocol}`)
      return { transactionId: 'transaction-1', expiresAt: 2_000 }
    }),
    quiesce: vi.fn(async () => { calls.push('quiesce'); return room() }),
    commitImport: vi.fn(async () => { calls.push('commit'); return receipt() }),
    abortImport: vi.fn(async () => { calls.push('abort') }),
    resume: vi.fn(async () => { calls.push('resume'); return room({ phase: 'active' }) }),
    activate: vi.fn(async () => { calls.push('activate') }),
    onRecoveryRequired: recovery,
    ...options.boundary,
  }
  return {
    calls,
    codec,
    boundary,
    recovery,
    coordinator: new XlsxCollaborativeImportCoordinator(exchange, boundary, () => 1_000),
  }
}

const replaceTarget = {
  mode: 'replace-room' as const,
  room: 'room-1',
  expectedArtifactVersion: 'version-7',
  expectedHeadSequence: 12,
}

describe('XlsxCollaborativeImportCoordinator', () => {
  it('authorizes, imports a server unit, commits it, then activates the durable room', async () => {
    const subject = harness()
    await expect(subject.coordinator.import({ source, target: { mode: 'new-room', requestedRoom: 'room-1' } }))
      .resolves.toEqual(receipt())

    expect(subject.calls).toEqual([
      `authorize:${XLSX_COLLABORATIVE_IMPORT_PROTOCOL}`, 'convert', 'commit', 'activate',
    ])
    expect(subject.boundary.commitImport).toHaveBeenCalledWith(expect.objectContaining({
      transactionId: 'transaction-1',
      importedUnit: expect.objectContaining({ unitId: 'imported-unit', revision: 'native-3' }),
      previous: undefined,
    }))
  })

  it('checks authorization before uploading workbook bytes', async () => {
    const subject = harness({ boundary: { authorizeAndBegin: vi.fn(async () => undefined) } })
    await expect(subject.coordinator.import({ source, target: { mode: 'new-room' } }))
      .rejects.toMatchObject({ code: 'permission' })
    expect(subject.codec.importServerUnit).not.toHaveBeenCalled()
    expect(subject.boundary.abortImport).not.toHaveBeenCalled()
  })

  it('quiesces and supplies an exact durable replacement head at commit', async () => {
    const subject = harness()
    await subject.coordinator.import({ source, target: replaceTarget })
    expect(subject.calls).toEqual([
      `authorize:${XLSX_COLLABORATIVE_IMPORT_PROTOCOL}`, 'convert', 'quiesce', 'commit', 'activate',
    ])
    expect(subject.boundary.commitImport).toHaveBeenCalledWith(expect.objectContaining({ previous: room(), target: replaceTarget }))
  })

  it.each([
    ['pending-outbound', room({ pendingOutbound: 1 })],
    ['unsaved-head', room({ savedSequence: 11 })],
    ['stale-head', room({ artifactVersion: 'version-6' })],
    ['stale-head', room({ headSequence: 13, savedSequence: 13 })],
    ['invalid-state', room({ phase: 'active' })],
  ] as const)('refuses an unsafe replacement state: %s', async (code, unsafe) => {
    const subject = harness({ boundary: { quiesce: vi.fn(async () => unsafe) } })
    await expect(subject.coordinator.import({ source, target: replaceTarget })).rejects.toMatchObject({ code })
    expect(subject.boundary.commitImport).not.toHaveBeenCalled()
    expect(subject.boundary.resume).toHaveBeenCalled()
    expect(subject.boundary.abortImport).toHaveBeenCalledWith(expect.objectContaining({ importedUnitId: 'imported-unit' }))
  })

  it('resumes the unchanged room and abandons staging when commit fails', async () => {
    const failure = new Error('STALE_BASE')
    const subject = harness({ boundary: { commitImport: vi.fn(async () => { throw failure }) } })
    await expect(subject.coordinator.import({ source, target: replaceTarget })).rejects.toBe(failure)
    expect(subject.calls).toEqual([
      `authorize:${XLSX_COLLABORATIVE_IMPORT_PROTOCOL}`, 'convert', 'quiesce', 'resume', 'abort',
    ])
  })

  it('propagates cancellation into conversion and abandons the transaction', async () => {
    let release!: () => void
    let entered!: () => void
    const started = new Promise<void>((resolve) => { entered = resolve })
    const codec: XlsxExchangeCodec<Snapshot> = {
      importServerUnit: vi.fn((_file, context) => new Promise<{ unitId: string }>((resolve) => {
        entered()
        release = () => resolve({ unitId: 'imported-unit' })
        expect(context.signal.aborted).toBe(false)
      })),
    }
    const subject = harness({ codec })
    const controller = new AbortController()
    const pending = subject.coordinator.import({ source, target: { mode: 'new-room' }, signal: controller.signal })
    await started
    controller.abort('dialog closed')
    release()
    await expect(pending).rejects.toMatchObject({ code: 'canceled' })
    expect(subject.boundary.commitImport).not.toHaveBeenCalled()
    expect(subject.boundary.abortImport).toHaveBeenCalled()
  })

  it('marks local activation failure as post-commit recovery, never rollback', async () => {
    const failure = new Error('editor reload failed')
    const subject = harness({ boundary: { activate: vi.fn(async () => { throw failure }) } })
    await expect(subject.coordinator.import({ source, target: { mode: 'new-room' } }))
      .rejects.toMatchObject({ name: 'XlsxCollaborativeImportRecoveryError', receipt: receipt() })
    expect(subject.recovery).toHaveBeenCalledWith({ receipt: receipt(), cause: failure })
    expect(subject.boundary.abortImport).not.toHaveBeenCalled()
    expect(subject.boundary.resume).not.toHaveBeenCalled()
  })

  it('treats a mismatched commit receipt as uncertain instead of issuing an unsafe abort', async () => {
    const subject = harness({ boundary: { commitImport: vi.fn(async () => receipt({ importedUnitId: 'other' })) } })
    await expect(subject.coordinator.import({ source, target: { mode: 'new-room' } }))
      .rejects.toMatchObject({ code: 'invalid-receipt' })
    expect(subject.boundary.activate).not.toHaveBeenCalled()
    expect(subject.boundary.abortImport).not.toHaveBeenCalled()
  })

  it('serializes imports while asynchronous authorization is in flight', async () => {
    let allow!: () => void
    let calls = 0
    const subject = harness({
      boundary: {
        authorizeAndBegin: vi.fn((): Promise<XlsxCollaborativeImportLease | undefined> => {
          calls += 1
          if (calls > 1) return Promise.resolve({ transactionId: 'transaction-1', expiresAt: 2_000 })
          return new Promise<XlsxCollaborativeImportLease>((resolve) => {
            allow = () => resolve({ transactionId: 'transaction-1', expiresAt: 2_000 })
          })
        }),
      },
    })
    const first = subject.coordinator.import({ source, target: { mode: 'new-room' } })
    await expect(subject.coordinator.import({ source, target: { mode: 'new-room' } }))
      .rejects.toBeInstanceOf(XlsxCollaborativeImportError)
    allow()
    await first
    await expect(subject.coordinator.import({ source, target: { mode: 'new-room' } })).resolves.toBeDefined()
  })

  it('refuses expired leases and malformed replacement targets', async () => {
    const expired = harness({ boundary: { authorizeAndBegin: vi.fn(async () => ({ transactionId: 'transaction-1', expiresAt: 999 })) } })
    await expect(expired.coordinator.import({ source, target: { mode: 'new-room' } }))
      .rejects.toMatchObject({ code: 'invalid-lease' })
    expect(expired.codec.importServerUnit).not.toHaveBeenCalled()
    expect(expired.boundary.abortImport).toHaveBeenCalled()

    const malformed = harness()
    await expect(malformed.coordinator.import({
      source,
      target: { mode: 'replace-room', room: '', expectedArtifactVersion: '', expectedHeadSequence: -1 },
    })).rejects.toMatchObject({ code: 'invalid-target' })
    expect(malformed.boundary.authorizeAndBegin).not.toHaveBeenCalled()
  })

  it('exports distinct typed errors for conflicts and committed recovery', () => {
    expect(new XlsxCollaborativeImportError('busy', 'busy')).toBeInstanceOf(Error)
    expect(new XlsxCollaborativeImportRecoveryError(receipt(), new Error('cause')).cause).toBeInstanceOf(Error)
  })
})
