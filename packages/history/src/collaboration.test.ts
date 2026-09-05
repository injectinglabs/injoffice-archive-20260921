import { describe, expect, it, vi } from 'vitest'
import {
  HistoryCollaborationConflictError,
  HistoryCollaborationCoordinator,
  type HistoryCollaborationBoundary,
  type HistoryCollaborationState,
} from './collaboration'
import type { HistoryVersionInfo } from './types'

const restored: HistoryVersionInfo = {
  id: 'v4', artifactId: 'book-1', sequence: 4, createdAt: 4_000, size: 20,
  contentType: 'xlsx', author: { id: 'ada', kind: 'user' }, reason: 'restore', sourceVersionId: 'v1',
  retention: { policyId: 'standard', expiresAt: null, legalHold: false },
}

function state(overrides: Partial<HistoryCollaborationState> = {}): HistoryCollaborationState {
  return {
    room: 'book-1', artifactVersionId: 'v3', logEpoch: 'epoch-v3', headSequence: 8,
    savedSequence: 8, pendingOutbound: 0, phase: 'quiesced', ...overrides,
  }
}

function harness(overrides: Partial<HistoryCollaborationBoundary> = {}) {
  const calls: string[] = []
  const recovery = vi.fn()
  const boundary: HistoryCollaborationBoundary = {
    authorizeRestore: vi.fn(async () => { calls.push('authorize'); return true }),
    quiesce: vi.fn(async () => { calls.push('quiesce'); return state() }),
    reloadAndReset: vi.fn(async () => { calls.push('reload') }),
    rejoin: vi.fn(async () => { calls.push('rejoin'); return state({ artifactVersionId: 'v4', logEpoch: 'epoch-v4', headSequence: 0, savedSequence: 0, phase: 'active' }) }),
    resume: vi.fn(async () => { calls.push('resume'); return state({ phase: 'active' }) }),
    onRecoveryRequired: recovery,
    ...overrides,
  }
  return { boundary, calls, coordinator: new HistoryCollaborationCoordinator('book-1', boundary), recovery }
}

describe('HistoryCollaborationCoordinator', () => {
  it('quiesces a saved epoch, resets to the durable version, and rejoins', async () => {
    const subject = harness()
    const lease = await subject.coordinator.prepare({ sourceVersionId: 'v1', expectedHeadVersionId: 'v3', signal: new AbortController().signal })
    await lease.activate(restored, { sourceVersionId: 'v1' })
    expect(subject.calls).toEqual(['authorize', 'quiesce', 'reload', 'rejoin'])
    expect(subject.boundary.reloadAndReset).toHaveBeenCalledWith({ room: 'book-1', version: restored, sourceVersionId: 'v1' })
    expect(subject.recovery).not.toHaveBeenCalled()
  })

  it.each([
    ['pending-outbound', state({ pendingOutbound: 1 })],
    ['unsaved-head', state({ savedSequence: 7 })],
    ['identity', state({ artifactVersionId: 'v2' })],
    ['phase', state({ phase: 'active' })],
  ] as const)('refuses unsafe pre-commit state: %s', async (code, unsafe) => {
    const subject = harness({ quiesce: async () => unsafe })
    await expect(subject.coordinator.prepare({ sourceVersionId: 'v1', expectedHeadVersionId: 'v3', signal: new AbortController().signal }))
      .rejects.toMatchObject({ code })
    expect(subject.boundary.resume).toHaveBeenCalled()
    expect(subject.boundary.reloadAndReset).not.toHaveBeenCalled()
  })

  it('checks permission before pausing the room', async () => {
    const subject = harness({ authorizeRestore: async () => false })
    await expect(subject.coordinator.prepare({ sourceVersionId: 'v1', signal: new AbortController().signal }))
      .rejects.toMatchObject({ code: 'permission' })
    expect(subject.boundary.quiesce).not.toHaveBeenCalled()
  })

  it('resumes the original epoch when durable creation is cancelled', async () => {
    const subject = harness()
    const lease = await subject.coordinator.prepare({ sourceVersionId: 'v1', signal: new AbortController().signal })
    await lease.cancel(new Error('version store failed'))
    expect(subject.calls).toEqual(['authorize', 'quiesce', 'resume'])
    await expect(lease.cancel(new Error('again'))).rejects.toMatchObject({ code: 'phase' })
  })

  it('reports a recovery boundary when reload or rejoin cannot activate', async () => {
    const failure = new Error('rejoin offline')
    const subject = harness({ rejoin: async () => { throw failure } })
    const lease = await subject.coordinator.prepare({ sourceVersionId: 'v1', signal: new AbortController().signal })
    await expect(lease.activate(restored, { sourceVersionId: 'v1' })).rejects.toBe(failure)
    expect(subject.recovery).toHaveBeenCalledWith({ room: 'book-1', version: restored, cause: failure })
  })

  it('serializes preparations and permits another after completion', async () => {
    const subject = harness()
    const first = await subject.coordinator.prepare({ sourceVersionId: 'v1', signal: new AbortController().signal })
    await expect(subject.coordinator.prepare({ sourceVersionId: 'v2', signal: new AbortController().signal }))
      .rejects.toBeInstanceOf(HistoryCollaborationConflictError)
    await first.cancel(new Error('cancel'))
    await expect(subject.coordinator.prepare({ sourceVersionId: 'v2', signal: new AbortController().signal })).resolves.toBeDefined()
  })

  it('serializes even while asynchronous authorization is in flight', async () => {
    let allow!: (value: boolean) => void
    const subject = harness({ authorizeRestore: () => new Promise<boolean>((resolve) => { allow = resolve }) })
    const first = subject.coordinator.prepare({ sourceVersionId: 'v1', signal: new AbortController().signal })
    await expect(subject.coordinator.prepare({ sourceVersionId: 'v2', signal: new AbortController().signal }))
      .rejects.toMatchObject({ code: 'busy' })
    allow(true)
    const lease = await first
    await lease.cancel(new Error('cancel'))
  })

  it('fails closed on invalid sequence metadata', async () => {
    const subject = harness({ quiesce: async () => state({ headSequence: -1 }) })
    await expect(subject.coordinator.prepare({ sourceVersionId: 'v1', signal: new AbortController().signal }))
      .rejects.toMatchObject({ code: 'invalid-sequence' })
  })
})
