import { describe, expect, it } from 'vitest'
import { groupThreads, newCommentEntry, newCommentId, openThreadCount, type CommentEntry } from './comments'

describe('groupThreads', () => {
  it('groups replies under their root, sorted by createdAt', () => {
    const entries: CommentEntry[] = [
      { id: 'a', author: 'Al', text: 'root A', createdAt: '2026-01-01T00:00:00Z' },
      { id: 'b', author: 'Bo', text: 'root B', createdAt: '2026-01-01T00:01:00Z' },
      { id: 'a2', parentId: 'a', author: 'Cy', text: 'reply to A, later', createdAt: '2026-01-01T00:03:00Z' },
      { id: 'a1', parentId: 'a', author: 'Dee', text: 'reply to A, earlier', createdAt: '2026-01-01T00:02:00Z' },
    ]
    const threads = groupThreads(entries)
    expect(threads).toHaveLength(2)
    const threadA = threads.find((t) => t.root.id === 'a')!
    expect(threadA.replies.map((r) => r.id)).toEqual(['a1', 'a2'])
    const threadB = threads.find((t) => t.root.id === 'b')!
    expect(threadB.replies).toEqual([])
  })

  it('a reply whose parent is missing from the entry list is simply dropped from any thread', () => {
    const entries: CommentEntry[] = [
      { id: 'orphan', parentId: 'ghost', author: 'X', text: 'orphaned reply', createdAt: '2026-01-01T00:00:00Z' },
    ]
    expect(groupThreads(entries)).toEqual([])
  })

  it('empty input yields no threads', () => {
    expect(groupThreads([])).toEqual([])
  })
})

describe('openThreadCount', () => {
  it('counts only unresolved thread roots', () => {
    const entries: CommentEntry[] = [
      { id: 'a', author: 'Al', text: 'open', createdAt: '2026-01-01T00:00:00Z' },
      { id: 'b', author: 'Bo', text: 'closed', createdAt: '2026-01-01T00:00:00Z', resolved: true },
    ]
    expect(openThreadCount(groupThreads(entries))).toBe(1)
  })

  it('zero threads yields zero', () => {
    expect(openThreadCount([])).toBe(0)
  })
})

describe('newCommentId / newCommentEntry', () => {
  it('generates distinct ids across calls', () => {
    const ids = new Set(Array.from({ length: 50 }, () => newCommentId()))
    expect(ids.size).toBe(50)
  })

  it('newCommentEntry stamps author/text/createdAt and an id', () => {
    const e = newCommentEntry('Nick', 'looks good')
    expect(e.author).toBe('Nick')
    expect(e.text).toBe('looks good')
    expect(e.id).toBeTruthy()
    expect(e.parentId).toBeUndefined()
    expect(() => new Date(e.createdAt).toISOString()).not.toThrow()
  })

  it('newCommentEntry can be built as a reply', () => {
    const e = newCommentEntry('Nick', 'agreed', 'root-1')
    expect(e.parentId).toBe('root-1')
  })
})
