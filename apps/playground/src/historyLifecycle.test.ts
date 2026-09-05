import { describe, expect, it } from 'vitest'
import { createPlaygroundHistory, playgroundHistoryAuthors } from './historyLifecycle'

describe('playground history lifecycle', () => {
  it('captures, lists newest-first, previews in isolation, and restores a new version', async () => {
    const { manager } = createPlaygroundHistory([
      { snapshot: 'one', contentType: 'text/plain', description: 'first' },
      { snapshot: 'two', contentType: 'text/plain', description: 'second' },
    ])
    const authors = playgroundHistoryAuthors()
    const page = await manager.listVersions()
    expect(page.versions.map((version) => version.id)).toEqual(['v2', 'v1'])
    expect(page.versions[0]?.sequence).toBeGreaterThan(page.versions[1]!.sequence)

    const captured = await manager.capture({
      snapshot: 'three',
      author: authors.agent,
      reason: 'agent-delivery',
      contentType: 'text/plain',
      expectedHeadVersionId: 'v2',
      description: 'agent save',
    })
    expect(captured.id).toBe('v3')
    expect(captured.author.kind).toBe('agent')

    const preview = await manager.loadPreview('v1')
    expect(preview.mode).toBe('isolated')
    expect(preview.snapshot).toBe('one')
    preview.snapshot = 'mutated'
    expect((await manager.loadPreview('v1')).snapshot).toBe('one')

    const restored = await manager.restore({ sourceVersionId: 'v1', author: authors.user, expectedHeadVersionId: 'v3' })
    expect(restored.reason).toBe('restore')
    expect(restored.sourceVersionId).toBe('v1')
    expect(restored.id).not.toBe('v1')
    expect((await manager.listVersions()).versions.map((version) => version.id)).toEqual(['v4', 'v3', 'v2', 'v1'])
  })

  it('filters before pagination and rejects a stale head', async () => {
    const { manager } = createPlaygroundHistory([
      { snapshot: 'one', contentType: 'text/plain', description: 'first' },
    ])
    const authors = playgroundHistoryAuthors()
    await manager.capture({ snapshot: 'agent', author: authors.agent, reason: 'agent-delivery', contentType: 'text/plain' })
    const agents = await manager.listVersions({ filter: { authorKinds: ['agent'] } })
    expect(agents.versions).toHaveLength(1)
    expect(agents.versions[0]?.author.kind).toBe('agent')
    await expect(manager.capture({
      snapshot: 'stale',
      author: authors.user,
      contentType: 'text/plain',
      expectedHeadVersionId: 'v1',
    })).rejects.toThrow(/expected head/)
  })
})
