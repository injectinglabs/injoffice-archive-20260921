import { describe, expect, it, vi } from 'vitest'
import { HistoryManager, HistoryProtocolError } from './manager'
import type { HistoryCreateRequest, HistoryHost, HistoryListRequest, HistoryLoadRequest, HistoryVersionInfo } from './types'
import { historyVersionMatches, validateHistoryVersionInfo } from './validation'

type Snapshot = { cells: string[][] }

const alice = { id: 'user-1', kind: 'user' as const, displayName: 'Alice' }
const robot = { id: 'agent-1', kind: 'agent' as const, displayName: 'Build agent' }

function version(overrides: Partial<HistoryVersionInfo> = {}): HistoryVersionInfo {
  return {
    id: 'v3',
    artifactId: 'book-1',
    sequence: 3,
    createdAt: 3_000,
    size: 42,
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    author: alice,
    reason: 'save',
    retention: { policyId: 'standard-30d', expiresAt: 2_595_000_000, legalHold: false },
    ...overrides,
  }
}

function host(overrides: Partial<HistoryHost<Snapshot>> = {}): HistoryHost<Snapshot> {
  return {
    async listVersions() { return { versions: [] } },
    async loadVersion(request) { return { version: version({ id: request.versionId }), snapshot: { cells: [['old']] } } },
    async createVersion(request) {
      return version({
        id: request.reason === 'restore' ? 'v4' : 'v3',
        sequence: request.reason === 'restore' ? 4 : 3,
        author: request.author,
        reason: request.reason,
        sourceVersionId: request.sourceVersionId,
      })
    },
    ...overrides,
  }
}

describe('HistoryManager listing and preview', () => {
  it('forwards validated pagination/filter options and returns newest-first metadata', async () => {
    let received: HistoryListRequest | undefined
    const versions = [
      version(),
      version({ id: 'v2', sequence: 2, createdAt: 2_000, author: robot, reason: 'agent-delivery' }),
    ]
    const manager = new HistoryManager(host({
      async listVersions(request) {
        received = request
        return { versions: [versions[1]!], nextCursor: 'next' }
      },
    }), 'book-1')

    const page = await manager.listVersions({ cursor: 'cursor', limit: 20, filter: { authorKinds: ['agent'], reasons: ['agent-delivery'] } })
    expect(received).toMatchObject({ artifactId: 'book-1', cursor: 'cursor', limit: 20 })
    expect(page).toEqual({ versions: [versions[1]], nextCursor: 'next' })
    page.versions[0]!.author.displayName = 'changed by caller'
    expect(versions[1]!.author.displayName).toBe('Build agent')
  })

  it('rejects invalid filters and host pages that are unfiltered or unordered', async () => {
    const unfiltered = new HistoryManager(host({ async listVersions() { return { versions: [version()] } } }), 'book-1')
    await expect(unfiltered.listVersions({ filter: { authorKinds: ['agent'] } })).rejects.toThrow('did not apply')
    await expect(unfiltered.listVersions({ limit: 0 })).rejects.toThrow('/limit')

    const unordered = new HistoryManager(host({
      async listVersions() { return { versions: [version({ id: 'v2', sequence: 2, createdAt: 2_000 }), version()] } },
    }), 'book-1')
    await expect(unordered.listVersions()).rejects.toThrow('newest first')
  })

  it('loads detached isolated previews without exposing the host snapshot', async () => {
    const stored = { cells: [['immutable']] }
    const manager = new HistoryManager(host({
      async loadVersion() { return { version: version({ id: 'v1', sequence: 1, createdAt: 1_000 }), snapshot: stored } },
    }), 'book-1')

    const preview = await manager.loadPreview('v1')
    expect(preview.mode).toBe('isolated')
    preview.snapshot.cells[0]![0] = 'preview mutation'
    preview.version.author.displayName = 'preview mutation'
    expect(stored.cells[0]![0]).toBe('immutable')
    expect(alice.displayName).toBe('Alice')
  })

  it('fails closed when a host returns metadata for another version', async () => {
    const manager = new HistoryManager(host({
      async loadVersion() { return { version: version({ id: 'wrong' }), snapshot: { cells: [] } } },
    }), 'book-1')
    await expect(manager.loadPreview('v1')).rejects.toBeInstanceOf(HistoryProtocolError)
  })
})

describe('HistoryManager capture and restore', () => {
  it('captures an author-attributed version with concurrency and retention requests', async () => {
    let received: HistoryCreateRequest<Snapshot> | undefined
    const manager = new HistoryManager(host({
      async createVersion(request) {
        received = request
        return version({ author: robot, reason: 'agent-delivery', retention: { policyId: 'audit', expiresAt: null, legalHold: true } })
      },
    }), 'book-1')
    const events = vi.fn()
    manager.onEvent(events)

    const created = await manager.capture({
      snapshot: { cells: [['new']] },
      author: robot,
      reason: 'agent-delivery',
      expectedHeadVersionId: 'v2',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      retention: { policyId: 'audit', legalHold: true },
    })

    expect(received).toMatchObject({ artifactId: 'book-1', expectedHeadVersionId: 'v2', author: robot, reason: 'agent-delivery', retention: { policyId: 'audit', legalHold: true } })
    expect(created.retention).toEqual({ policyId: 'audit', expiresAt: null, legalHold: true })
    expect(events).toHaveBeenCalledWith(expect.objectContaining({ name: 'captured', version: created }))
  })

  it('restores by loading an old snapshot and creating a new version with lineage', async () => {
    const source = { cells: [['old']] }
    let received: HistoryCreateRequest<Snapshot> | undefined
    const manager = new HistoryManager(host({
      async loadVersion(request: HistoryLoadRequest) {
        return { version: version({ id: request.versionId, sequence: 1, createdAt: 1_000 }), snapshot: source }
      },
      async createVersion(request) {
        received = request
        request.snapshot.cells[0]![0] = 'host mutation'
        return version({ id: 'v4', sequence: 4, createdAt: 4_000, author: request.author, reason: 'restore', sourceVersionId: 'v1' })
      },
    }), 'book-1')
    const events = vi.fn()
    manager.onEvent(events)

    const restored = await manager.restore({ sourceVersionId: 'v1', author: alice, expectedHeadVersionId: 'v3' })
    expect(received).toMatchObject({ artifactId: 'book-1', sourceVersionId: 'v1', reason: 'restore', expectedHeadVersionId: 'v3' })
    expect(received?.contentType).toContain('spreadsheetml')
    expect(source.cells[0]![0]).toBe('old')
    expect(restored).toMatchObject({ id: 'v4', reason: 'restore', sourceVersionId: 'v1' })
    expect(events).toHaveBeenCalledWith(expect.objectContaining({ name: 'restored', sourceVersionId: 'v1' }))
  })

  it('rejects a host that overwrites the restored source version', async () => {
    const manager = new HistoryManager(host({
      async createVersion(request) { return version({ id: 'v1', sequence: 1, createdAt: 1_000, reason: request.reason, sourceVersionId: request.sourceVersionId }) },
    }), 'book-1')
    await expect(manager.restore({ sourceVersionId: 'v1', author: alice })).rejects.toThrow('new immutable version')
  })

  it('does not create a durable restore when cancellation arrives after source loading began', async () => {
    let release!: () => void
    const pending = new Promise<void>((resolve) => { release = resolve })
    const createVersion = vi.fn(host().createVersion)
    const manager = new HistoryManager(host({
      async loadVersion(request) {
        await pending
        return { version: version({ id: request.versionId }), snapshot: { cells: [['old']] } }
      },
      createVersion,
    }), 'book-1')
    const abort = new AbortController()
    const restoring = manager.restore({ sourceVersionId: 'v1', author: alice, signal: abort.signal })
    abort.abort()
    release()
    await expect(restoring).rejects.toMatchObject({ name: 'AbortError' })
    expect(createVersion).not.toHaveBeenCalled()
  })

  it('fails closed when storage changes attribution', async () => {
    const manager = new HistoryManager(host({
      async createVersion() { return version({ author: robot }) },
    }), 'book-1')
    await expect(manager.capture({ snapshot: { cells: [] }, author: alice, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })).rejects.toThrow('author')
  })

  it('rejects non-cloneable snapshots before invoking storage', async () => {
    const createVersion = vi.fn(host().createVersion)
    const manager = new HistoryManager(host({ createVersion }), 'book-1')
    await expect(manager.capture({ snapshot: { cells: [] }, author: alice, contentType: 'xlsx', description: 'x'.repeat(2049) })).rejects.toThrow('description')
    await expect(manager.capture({ snapshot: { cells: [] }, author: alice, contentType: 'xlsx', retention: { legalHold: 'yes' as unknown as boolean } })).rejects.toThrow('legalHold')
    await expect(manager.capture({ snapshot: { cells: [], callback: () => undefined } as unknown as Snapshot, author: alice, contentType: 'xlsx' })).rejects.toThrow('structured-cloneable')
    expect(createVersion).not.toHaveBeenCalled()
  })
})

describe('history validation', () => {
  it('validates restore lineage, retention, and filter matching', () => {
    expect(validateHistoryVersionInfo(version({ reason: 'restore', sourceVersionId: undefined })).ok).toBe(false)
    expect(validateHistoryVersionInfo(version({ retention: { policyId: 'x', expiresAt: 1, legalHold: false } })).ok).toBe(false)
    expect(historyVersionMatches(version({ author: robot }), { authorIds: ['agent-1'], createdAfter: 2_999 })).toBe(true)
    expect(historyVersionMatches(version(), { reasons: ['import'] })).toBe(false)
  })
})
