import { describe, expect, it, vi } from 'vitest'
import type { XlsxCollaborationRoomState } from './collaboration'
import {
  XLSX_COLLABORATIVE_EXPORT_PROTOCOL,
  XlsxCollaborativeExportCoordinator,
  XlsxCollaborativeExportError,
  XlsxCollaborativeExportRecoveryError,
  type XlsxCollaborativeExportBeginResult,
  type XlsxCollaborativeExportBoundary,
  type XlsxCollaborativeExportFileBinding,
  type XlsxCollaborativeExportLease,
  type XlsxCollaborativeExportReceipt,
} from './exportCollaboration'
import { sha256HexBytes } from './exportHash'
import { XlsxExchangeManager } from './manager'
import type { XlsxExchangeCodec, XlsxExchangeContext, XlsxExchangeStorage } from './types'

interface Snapshot { readonly cells: readonly string[] }
const bytes = new Uint8Array([7, 8, 9])

function state(phase: XlsxCollaborationRoomState['phase'] = 'quiesced', overrides: Partial<XlsxCollaborationRoomState> = {}): XlsxCollaborationRoomState {
  return {
    room: 'room-1', unitId: 'unit-7', artifactVersion: 'artifact-7', logEpoch: 'epoch-7',
    headSequence: 12, savedSequence: 12, pendingOutbound: 0, phase, ...overrides,
  }
}

function lease(requestId = 'request-1', overrides: Partial<XlsxCollaborativeExportLease> = {}): XlsxCollaborativeExportLease {
  return {
    requestId,
    transactionId: 'transaction-1',
    expiresAt: 2_000,
    state: state(),
    snapshot: {
      room: 'room-1', sourceUnitId: 'unit-7', exportUnitId: 'export-unit-7', artifactVersion: 'artifact-7',
      logEpoch: 'epoch-7', headSequence: 12, revision: 'snapshot-12',
    },
    ...overrides,
  }
}

function binding(overrides: Partial<XlsxCollaborativeExportFileBinding> = {}): XlsxCollaborativeExportFileBinding {
  return {
    revision: 'snapshot-12', byteLength: bytes.byteLength,
    sha256: `sha256:${sha256HexBytes(bytes)}`, location: 'download://report.xlsx', ...overrides,
  }
}

function receipt(requestId = 'request-1', overrides: Partial<XlsxCollaborativeExportReceipt> = {}): XlsxCollaborativeExportReceipt {
  return {
    requestId,
    transactionId: 'transaction-1',
    snapshot: lease(requestId).snapshot,
    file: binding(),
    state: state('active'),
    ...overrides,
  }
}

const request = {
  room: 'room-1', expectedArtifactVersion: 'artifact-7', expectedLogEpoch: 'epoch-7',
  expectedHeadSequence: 12, destination: 'downloads',
}

function harness(options: {
  codec?: XlsxExchangeCodec<Snapshot>
  boundary?: Partial<XlsxCollaborativeExportBoundary<string>>
  maxOutputBytes?: number
} = {}) {
  const calls: string[] = []
  const codec: XlsxExchangeCodec<Snapshot> = options.codec ?? {
    exportServerUnit: vi.fn(async (unitId: string) => {
      calls.push(`export:${unitId}`)
      return { bytes: bytes.slice(), name: 'report.xlsx', mediaType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', revision: 'snapshot-12' }
    }),
  }
  const storage = {
    read: vi.fn(),
    saveAs: vi.fn(async (_file, destination: string | undefined) => {
      calls.push(`save:${destination}`)
      return { location: 'download://report.xlsx' }
    }),
  } satisfies XlsxExchangeStorage<string, string>
  const exchange = new XlsxExchangeManager<Snapshot, string, string>({ codec, storage, maxOutputBytes: options.maxOutputBytes ?? 1024 })
  const recovery = vi.fn()
  const boundary: XlsxCollaborativeExportBoundary<string> = {
    authorizeAndCapture: vi.fn(async ({ requestId, protocol }): Promise<XlsxCollaborativeExportBeginResult> => {
      calls.push(`begin:${protocol}`)
      return { status: 'granted', lease: lease(requestId) }
    }),
    completeExport: vi.fn(async ({ requestId, file }) => {
      calls.push('complete')
      return receipt(requestId, { file })
    }),
    resolveExport: vi.fn(async ({ requestId }) => ({ status: 'aborted' as const, requestId, state: state('active') })),
    ...options.boundary,
  }
  let id = 0
  const coordinator = new XlsxCollaborativeExportCoordinator(exchange, boundary, {
    idFactory: () => `request-${++id}`,
    now: () => 1_000,
    onRecoveryRequired: recovery,
  })
  return { boundary, calls, codec, coordinator, exchange, recovery, storage }
}

describe('XlsxCollaborativeExportCoordinator', () => {
  it('exports only the immutable server capture and commits an exact byte/head binding', async () => {
    const subject = harness()
    const result = await subject.coordinator.export(request)

    expect(subject.calls).toEqual([
      `begin:${XLSX_COLLABORATIVE_EXPORT_PROTOCOL}`, 'export:export-unit-7', 'save:downloads', 'complete',
    ])
    expect(result.receipt).toEqual(receipt())
    expect(result.export.file.bytes).toEqual(bytes)
    expect(subject.boundary.completeExport).toHaveBeenCalledWith(expect.objectContaining({
      transactionId: 'transaction-1', snapshot: lease().snapshot, file: binding(),
    }))
    result.export.file.bytes.fill(0)
    expect(bytes).toEqual(new Uint8Array([7, 8, 9]))
    expect(subject.coordinator.state).toEqual({ busy: false, blocked: false, recovery: undefined })
  })

  it.each([
    ['denied', 'permission'], ['busy', 'busy'], ['stale-head', 'stale-head'],
    ['pending-outbound', 'pending-outbound'], ['unsaved-head', 'unsaved-head'], ['invalid-state', 'invalid-state'],
  ] as const)('fails a definitive begin rejection without exporting: %s', async (status, code) => {
    const subject = harness({ boundary: { authorizeAndCapture: vi.fn(async () => ({ status })) } })
    await expect(subject.coordinator.export(request)).rejects.toMatchObject({ code })
    expect(subject.codec.exportServerUnit).not.toHaveBeenCalled()
    expect(subject.coordinator.state.blocked).toBe(false)
  })

  it('refuses a grant that is not the exact quiesced saved requested head', async () => {
    const unsafe = lease('request-1', { state: state('quiesced', { pendingOutbound: 1 }) })
    const subject = harness({ boundary: { authorizeAndCapture: vi.fn(async () => ({ status: 'granted' as const, lease: unsafe })) } })
    await expect(subject.coordinator.export(request)).rejects.toBeInstanceOf(XlsxCollaborativeExportRecoveryError)
    expect(subject.codec.exportServerUnit).not.toHaveBeenCalled()
    expect(subject.coordinator.state.blocked).toBe(true)
  })

  it('blocks ambiguous begin failure and clears only after exact aborted resync', async () => {
    const subject = harness({ boundary: { authorizeAndCapture: vi.fn(async () => { throw new Error('response lost') }) } })
    await expect(subject.coordinator.export(request)).rejects.toMatchObject({ code: 'uncertain-outcome' })
    await expect(subject.coordinator.export(request)).rejects.toMatchObject({ code: 'blocked' })
    expect(subject.recovery).toHaveBeenCalledWith(expect.objectContaining({ requestId: 'request-1', room: 'room-1' }), expect.any(Error))
    await expect(subject.coordinator.resync()).resolves.toEqual({ status: 'aborted', state: state('active') })
    expect(subject.coordinator.state.blocked).toBe(false)
  })

  it('keeps a saved export across uncertain completion and recovers the committed receipt', async () => {
    const subject = harness({ boundary: {
      completeExport: vi.fn(async () => { throw new Error('commit response lost') }),
      resolveExport: vi.fn(async () => ({ status: 'committed' as const, receipt: receipt() })),
    } })
    await expect(subject.coordinator.export(request)).rejects.toMatchObject({ code: 'uncertain-outcome' })
    expect(subject.storage.saveAs).toHaveBeenCalledOnce()
    const resolved = await subject.coordinator.resync()
    expect(resolved).toMatchObject({ status: 'committed', receipt: receipt(), export: { file: { bytes } } })
    expect(subject.coordinator.state.blocked).toBe(false)
  })

  it('keeps invalid resync blocked and never accepts a different head', async () => {
    const subject = harness({ boundary: {
      authorizeAndCapture: vi.fn(async () => { throw new Error('response lost') }),
      resolveExport: vi.fn(async ({ requestId }) => ({ status: 'aborted' as const, requestId, state: state('active', { headSequence: 13, savedSequence: 13 }) })),
    } })
    await expect(subject.coordinator.export(request)).rejects.toBeInstanceOf(XlsxCollaborativeExportRecoveryError)
    await expect(subject.coordinator.resync()).rejects.toBeInstanceOf(XlsxCollaborativeExportRecoveryError)
    expect(subject.coordinator.state.blocked).toBe(true)
  })

  it('keeps resync blocked when a recovered receipt does not bind its file revision to the snapshot', async () => {
    const subject = harness({ boundary: {
      authorizeAndCapture: vi.fn(async () => { throw new Error('response lost') }),
      resolveExport: vi.fn(async () => ({
        status: 'committed' as const,
        receipt: receipt('request-1', { file: binding({ revision: 'different-revision' }) }),
      })),
    } })
    await expect(subject.coordinator.export(request)).rejects.toBeInstanceOf(XlsxCollaborativeExportRecoveryError)
    await expect(subject.coordinator.resync()).rejects.toBeInstanceOf(XlsxCollaborativeExportRecoveryError)
    expect(subject.coordinator.state.blocked).toBe(true)
  })

  it('blocks exporter and cancellation faults because destination outcome can be uncertain', async () => {
    const failed = harness({ codec: { exportServerUnit: vi.fn(async () => { throw new Error('native service disconnected') }) } })
    await expect(failed.coordinator.export(request)).rejects.toMatchObject({ code: 'uncertain-outcome' })
    expect(failed.coordinator.state.blocked).toBe(true)

    let release!: () => void
    let started!: () => void
    const entered = new Promise<void>((resolve) => { started = resolve })
    const canceled = harness({ codec: {
      exportServerUnit: vi.fn((_unitId: string, context: XlsxExchangeContext) => new Promise<{ bytes: Uint8Array; mediaType: string; revision: string }>((resolve) => {
        started()
        release = () => resolve({ bytes: bytes.slice(), mediaType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', revision: 'snapshot-12' })
        expect(context.signal.aborted).toBe(false)
      })),
    } })
    const controller = new AbortController()
    const pending = canceled.coordinator.export({ ...request, signal: controller.signal })
    await entered
    controller.abort('dialog closed')
    release()
    await expect(pending).rejects.toMatchObject({ code: 'uncertain-outcome' })
    expect(canceled.coordinator.state.blocked).toBe(true)
  })

  it('blocks a failed save and isolates recovery observers', async () => {
    const subject = harness()
    subject.storage.saveAs.mockRejectedValue(new Error('destination response lost'))
    subject.recovery.mockImplementation(() => { throw new Error('observer failed') })

    await expect(subject.coordinator.export(request)).rejects.toBeInstanceOf(XlsxCollaborativeExportRecoveryError)
    expect(subject.recovery).toHaveBeenCalledOnce()
    expect(subject.boundary.completeExport).not.toHaveBeenCalled()
    expect(subject.coordinator.state.blocked).toBe(true)
  })

  it('rejects revision mismatches and bounds bytes before saveAs', async () => {
    const mismatch = harness({ codec: { exportServerUnit: vi.fn(async () => ({ bytes: bytes.slice(), revision: 'wrong' })) } })
    await expect(mismatch.coordinator.export(request)).rejects.toMatchObject({ code: 'uncertain-outcome' })
    expect(mismatch.storage.saveAs).not.toHaveBeenCalled()
    expect(mismatch.coordinator.state.blocked).toBe(true)

    const oversized = harness({
      maxOutputBytes: 2,
      codec: { exportServerUnit: vi.fn(async () => ({ bytes: bytes.slice(), revision: 'snapshot-12' })) },
    })
    await expect(oversized.coordinator.export(request)).rejects.toMatchObject({ code: 'uncertain-outcome' })
    expect(oversized.storage.saveAs).not.toHaveBeenCalled()
  })

  it('allows only one server lease across two coordinators', async () => {
    let held = false
    let release!: () => void
    const wait = new Promise<void>((resolve) => { release = resolve })
    const shared = harness({ boundary: {
      authorizeAndCapture: vi.fn(async ({ requestId }): Promise<XlsxCollaborativeExportBeginResult> => {
        if (held) return { status: 'busy' }
        held = true
        return { status: 'granted', lease: lease(requestId, { transactionId: `tx-${requestId}` }) }
      }),
      completeExport: vi.fn(async ({ requestId, transactionId, file }) => {
        await wait
        held = false
        return receipt(requestId, { transactionId, file })
      }),
    } })
    let id = 1
    const second = new XlsxCollaborativeExportCoordinator(shared.exchange, shared.boundary, { idFactory: () => `request-${++id}`, now: () => 1_000 })
    const firstPending = shared.coordinator.export(request)
    await Promise.resolve(); await Promise.resolve()
    await expect(second.export(request)).rejects.toMatchObject({ code: 'busy' })
    release()
    await expect(firstPending).resolves.toBeDefined()
    await expect(second.export(request)).resolves.toBeDefined()
  })

  it('validates requests and serializes one coordinator before contacting the server', async () => {
    let release!: () => void
    const subject = harness({ boundary: {
      authorizeAndCapture: vi.fn(({ requestId }) => new Promise<XlsxCollaborativeExportBeginResult>((resolve) => {
        release = () => resolve({ status: 'granted', lease: lease(requestId) })
      })),
    } })
    const first = subject.coordinator.export(request)
    await Promise.resolve()
    await expect(subject.coordinator.export(request)).rejects.toBeInstanceOf(XlsxCollaborativeExportError)
    release()
    await first
    await expect(subject.coordinator.export({ ...request, expectedHeadSequence: -1 })).rejects.toMatchObject({ code: 'invalid-request' })
    await expect(subject.coordinator.export({ ...request, destination: 'x'.repeat(70_000) })).rejects.toMatchObject({ code: 'invalid-request' })
    expect(subject.boundary.authorizeAndCapture).toHaveBeenCalledTimes(1)
  })
})

describe('collaborative export SHA-256', () => {
  it('matches standard binary vectors', () => {
    expect(sha256HexBytes(new Uint8Array())).toBe('e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855')
    expect(sha256HexBytes(new TextEncoder().encode('abc'))).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
    expect(sha256HexBytes(new Uint8Array([0, 255, 1]))).toBe('47ffa3ea45a70b8a41c2c0825df323c00a8b7a01c1ea06083cc41dddcc001123')
  })
})
