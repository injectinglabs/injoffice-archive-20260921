import { readFileSync } from 'node:fs'
import { EditorState } from 'prosemirror-state'
import { collab, receiveTransaction, sendableSteps } from 'prosemirror-collab'
import { describe, expect, it } from 'vitest'
import { opsFromSteps, seedDocsCollabDoc, stepsFromOps } from './collab/docsEditor'

describe('playground document collab proof', () => {
  it('keeps CollabDocsPlaceholder as a live ProseMirror collab editor', () => {
    const source = readFileSync(new URL('./collab/docs.tsx', import.meta.url), 'utf8')
    expect(source).toContain('export function CollabDocsPlaceholder')
    expect(source).toContain('DocPresenceManager')
    expect(source).toContain('startSync')
    expect(source).toContain('sendableSteps')
    expect(source).toContain('receiveTransaction')
    expect(source).toContain('publish({')
    expect(source).toContain('mintCollabArtifact')
    expect(source).toContain('fetchSampleWorkbook')
    expect(source).not.toMatch(/nativeTransactionAdapter|adaptNativeDocxProseMirror/)
    expect(source).not.toContain('@injoffice/docs')
  })

  it('round-trips ProseMirror steps through sendableSteps and receiveTransaction', () => {
    let author = EditorState.create({
      doc: seedDocsCollabDoc(),
      plugins: [collab({ version: 0, clientID: 'a' })],
    })
    let follower = EditorState.create({
      doc: seedDocsCollabDoc(),
      plugins: [collab({ version: 0, clientID: 'b' })],
    })
    author = author.apply(author.tr.insertText('Hi ', 1))
    const sendable = sendableSteps(author)
    expect(sendable).not.toBeNull()
    const ops = opsFromSteps(sendable!.steps)
    const remote = stepsFromOps(ops)
    follower = follower.apply(receiveTransaction(follower, remote, remote.map(() => 'a')))
    author = author.apply(receiveTransaction(author, remote, remote.map(() => 'a')))
    expect(author.doc.textContent).toBe(follower.doc.textContent)
    expect(author.doc.textContent.startsWith('Hi ')).toBe(true)
    expect(sendableSteps(author)).toBeNull()
  })
})
