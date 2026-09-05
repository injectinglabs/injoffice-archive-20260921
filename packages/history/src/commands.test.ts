import { describe, expect, it, vi } from 'vitest'
import {
  HistoryActivationError,
  HistoryCommandBusyError,
  HistoryCommandController,
  HistoryPreparationReleaseError,
  isHistoryAuthor,
  type HistoryWorkspace,
} from './commands'
import { HistoryManager } from './manager'
import type { HistoryHost, HistoryVersionInfo } from './types'

type Snapshot = { cells: string[][] }
const author = { id: 'user-1', kind: 'user' as const, displayName: 'Ada' }

function version(overrides: Partial<HistoryVersionInfo> = {}): HistoryVersionInfo {
  return {
    id: 'v3', artifactId: 'book-1', sequence: 3, createdAt: 3_000, size: 20,
    contentType: 'xlsx', author, reason: 'save',
    retention: { policyId: 'standard', expiresAt: null, legalHold: false },
    ...overrides,
  }
}

function subject(overrides: Partial<HistoryHost<Snapshot>> = {}, workspaceOverrides: Partial<HistoryWorkspace<Snapshot>> = {}) {
  const host: HistoryHost<Snapshot> = {
    async listVersions() { return { versions: [version()] } },
    async loadVersion({ versionId }) { return { version: version({ id: versionId }), snapshot: { cells: [['old']] } } },
    async createVersion(request) {
      return version({
        id: request.reason === 'restore' ? 'v4' : 'v3',
        sequence: request.reason === 'restore' ? 4 : 3,
        author: request.author,
        reason: request.reason,
        sourceVersionId: request.sourceVersionId,
        contentType: request.contentType,
      })
    },
    ...overrides,
  }
  const workspace: HistoryWorkspace<Snapshot> = {
    captureSnapshot: vi.fn(async () => ({ snapshot: { cells: [['live']] }, contentType: 'xlsx' })),
    showIsolatedPreview: vi.fn(),
    activateRestoredVersion: vi.fn(),
    ...workspaceOverrides,
  }
  const controller = new HistoryCommandController(new HistoryManager(host, 'book-1'), workspace)
  return { controller, workspace }
}

describe('HistoryCommandController', () => {
  it('coordinates list, isolated preview, capture, restore activation, and state', async () => {
    const { controller, workspace } = subject()
    const states = vi.fn()
    controller.onState(states)

    await expect(controller.list()).resolves.toMatchObject({ versions: [{ id: 'v3' }] })
    const preview = await controller.preview('v1')
    expect(preview).toMatchObject({ mode: 'isolated', version: { id: 'v1' } })
    expect(workspace.showIsolatedPreview).toHaveBeenCalledWith(preview, expect.objectContaining({ signal: expect.any(AbortSignal) }))
    await expect(controller.capture({ author, expectedHeadVersionId: 'v2' })).resolves.toMatchObject({ id: 'v3' })
    const restored = await controller.restore({ sourceVersionId: 'v1', author, expectedHeadVersionId: 'v3' })
    expect(restored).toMatchObject({ id: 'v4', sourceVersionId: 'v1' })
    expect(workspace.activateRestoredVersion).toHaveBeenCalledWith(restored, { sourceVersionId: 'v1' })
    expect(controller.getState()).toMatchObject({ operation: 'restore', status: 'succeeded' })
    expect(states).toHaveBeenCalledWith(expect.objectContaining({ operation: 'preview', status: 'running' }))
  })

  it('serializes jobs and cancels the active request', async () => {
    let release!: () => void
    const pending = new Promise<void>((resolve) => { release = resolve })
    const { controller } = subject({ async listVersions(request) {
      await pending
      if (request.signal?.aborted) throw new DOMException('cancelled', 'AbortError')
      return { versions: [] }
    } })

    const running = controller.list()
    await expect(controller.preview('v1')).rejects.toBeInstanceOf(HistoryCommandBusyError)
    expect(controller.cancel()).toBe(true)
    release()
    await expect(running).rejects.toMatchObject({ name: 'AbortError' })
    expect(controller.getState()).toMatchObject({ operation: 'list', status: 'cancelled' })
    expect(controller.cancel()).toBe(false)
  })

  it('reports the durable version when live activation fails', async () => {
    const { controller } = subject({}, { activateRestoredVersion: async () => { throw new Error('reload failed') } })
    const result = controller.restore({ sourceVersionId: 'v1', author })
    await expect(result).rejects.toBeInstanceOf(HistoryActivationError)
    await expect(result).rejects.toMatchObject({ version: { id: 'v4' }, cause: expect.any(Error) })
    expect(controller.getState()).toMatchObject({ status: 'failed', error: expect.stringContaining('v4') })
  })

  it('starts a replacement after cancellation without stale state clobbering', async () => {
    let release!: () => void
    const pending = new Promise<void>((resolve) => { release = resolve })
    const { controller } = subject({ async listVersions() { await pending; return { versions: [] } } })
    const first = controller.list()
    controller.cancel()
    await expect(controller.preview('v1')).resolves.toMatchObject({ version: { id: 'v1' } })
    release()
    await expect(first).rejects.toMatchObject({ name: 'AbortError' })
    expect(controller.getState()).toMatchObject({ operation: 'preview', status: 'succeeded' })
  })

  it('prepares collaboration before restore and activates through its lease', async () => {
    const order: string[] = []
    const activate = vi.fn(async () => { order.push('activate') })
    const cancel = vi.fn()
    const { controller, workspace } = subject({
      async loadVersion(request) { order.push('load'); return { version: version({ id: request.versionId }), snapshot: { cells: [] } } },
      async createVersion(request) { order.push('create'); return version({ id: 'v4', sequence: 4, reason: 'restore', sourceVersionId: request.sourceVersionId }) },
    }, { prepareRestore: vi.fn(async () => { order.push('prepare'); return { activate, cancel } }) })

    const restored = await controller.restore({ sourceVersionId: 'v1', expectedHeadVersionId: 'v3', author })
    expect(order).toEqual(['prepare', 'load', 'create', 'activate'])
    expect(activate).toHaveBeenCalledWith(restored, { sourceVersionId: 'v1' })
    expect(cancel).not.toHaveBeenCalled()
    expect(workspace.activateRestoredVersion).not.toHaveBeenCalled()
  })

  it('resumes a prepared workspace when durable restore fails', async () => {
    const restoreError = new Error('storage conflict')
    const cancel = vi.fn()
    const { controller } = subject({ async createVersion() { throw restoreError } }, {
      prepareRestore: async () => ({ activate: vi.fn(), cancel }),
    })
    await expect(controller.restore({ sourceVersionId: 'v1', author })).rejects.toBe(restoreError)
    expect(cancel).toHaveBeenCalledWith(restoreError)
  })

  it('reports both restore and resume failures', async () => {
    const restoreError = new Error('storage conflict')
    const releaseError = new Error('resume failed')
    const { controller } = subject({ async createVersion() { throw restoreError } }, {
      prepareRestore: async () => ({ activate: vi.fn(), cancel: async () => { throw releaseError } }),
    })
    await expect(controller.restore({ sourceVersionId: 'v1', author })).rejects.toMatchObject({
      name: HistoryPreparationReleaseError.name,
      restoreError,
      releaseError,
    })
  })
})

describe('history browser guards', () => {
  it('recognizes stable attributed authors', () => {
    expect(isHistoryAuthor(author)).toBe(true)
    expect(isHistoryAuthor({ id: '', kind: 'user' })).toBe(false)
    expect(isHistoryAuthor({ id: 'x', kind: 'bot' })).toBe(false)
  })
})
