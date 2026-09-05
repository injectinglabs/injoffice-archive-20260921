// D7 threaded comments — pure data model + small pure helpers. The actual
// anchoring/remapping through concurrent edits is NOT reimplemented here:
// comments ride the collab op log as ordinary ProseMirror marks (see
// packages/collab's sync.ts/transform.ts docs and the website-local
// docComments.ts thin shell), so position tracking through edits comes for
// free from prosemirror-transform's own step mapping — the same mechanism
// that already keeps peer cursors and the undo stack correct. There is no
// injoffice-specific structural transform to write for marks the way Sheets
// needed one for cell ranges (Sheets isn't ProseMirror-native; Docs already
// is). What IS worth a pure, tested module is the comment data shape itself
// and the bookkeeping around it (id generation, thread grouping, resolved
// counts) — the same split this project uses everywhere else (pure
// core / thin DOM shell).

import { compareNativeCodeUnits } from './nativeDeterminism.js'

/** One comment (a thread root, or a reply when parentId is set). */
export interface CommentEntry {
  id: string
  /** the thread root's id this entry replies to; unset for the root itself. */
  parentId?: string
  author: string
  text: string
  /** ISO 8601. */
  createdAt: string
  /** only meaningful on a thread root; a reply's resolved state is ignored. */
  resolved?: boolean
}

/** A thread root plus its replies, in reply order. */
export interface CommentThread {
  root: CommentEntry
  replies: CommentEntry[]
}

/** Group a flat entry list (as stored on marks) into threads, root-first. */
export function groupThreads(entries: CommentEntry[]): CommentThread[] {
  const roots = entries.filter((e) => !e.parentId)
  const byParent = new Map<string, CommentEntry[]>()
  for (const e of entries) {
    if (!e.parentId) continue
    const list = byParent.get(e.parentId) ?? []
    list.push(e)
    byParent.set(e.parentId, list)
  }
  return roots.map((root) => ({
    root,
    replies: (byParent.get(root.id) ?? []).sort((a, b) => compareNativeCodeUnits(a.createdAt, b.createdAt)),
  }))
}

/** Unresolved thread count — the number the rail's badge shows. */
export function openThreadCount(threads: CommentThread[]): number {
  return threads.filter((t) => !t.root.resolved).length
}

/**
 * A short, sufficiently-unique id for a new comment mark. Not a UUID (no
 * crypto dependency needed for something scoped to one document's comment
 * count) — timestamp + a short random suffix is enough to avoid collisions
 * within a single session, and ids only need to be unique within one
 * document.
 */
export function newCommentId(): string {
  const rand = Math.random().toString(36).slice(2, 8)
  return `c${Date.now().toString(36)}${rand}`
}

export function newCommentEntry(author: string, text: string, parentId?: string): CommentEntry {
  return {
    id: newCommentId(),
    parentId,
    author,
    text,
    createdAt: new Date().toISOString(),
  }
}
