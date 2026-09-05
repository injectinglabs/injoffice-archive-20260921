import { useEffect, useRef, useState } from 'react'
import { collab, receiveTransaction, sendableSteps } from 'prosemirror-collab'
import { EditorState } from 'prosemirror-state'
import { EditorView } from 'prosemirror-view'
import {
  DocPresenceManager,
  type DocSelection,
  type DocSyncEngine,
} from '../../../packages/collab/src/index.js'
import type { BrowserCollabHub } from './browserCollabTransport'
import { opsFromSteps, seedDocsCollabDoc, stepsFromOps } from './collab/docsEditor'
import { SIM_EDITORS, SIM_ROOMS, SimEditorFrame, type SimEditorProfile, type SimMetrics } from './collabSimChrome'

function caretLabel(selection: DocSelection | null | undefined): string {
  if (!selection) return 'no caret'
  return `{from: ${selection.from}, to: ${selection.to}}`
}

function SimulatedDocsEditor({ hub, profile }: { hub: BrowserCollabHub; profile: SimEditorProfile }) {
  const editorRef = useRef<HTMLDivElement | null>(null)
  const viewRef = useRef<EditorView | null>(null)
  const engineRef = useRef<DocSyncEngine | null>(null)
  const collabClientIdRef = useRef(`docs-${profile.id}-${crypto.randomUUID()}`)
  const [status, setStatus] = useState<'starting' | 'live' | 'error'>('starting')
  const [metrics, setMetrics] = useState<SimMetrics>({ peers: 0, applied: 0, pending: 0 })
  const [presence, setPresence] = useState<DocPresenceManager | null>(null)
  const [remoteCaret, setRemoteCaret] = useState('no remote caret')

  useEffect(() => {
    const host = editorRef.current
    if (!host) return
    let disposed = false
    let stopEditor: (() => void) | null = null
    const clientID = collabClientIdRef.current
    const presenceRef = { current: null as DocPresenceManager | null }
    const view = new EditorView(host, {
      state: EditorState.create({
        doc: seedDocsCollabDoc(),
        plugins: [collab({ version: 0, clientID })],
      }),
      dispatchTransaction(tr) {
        const next = view.state.apply(tr)
        view.updateState(next)
        if (engineRef.current && sendableSteps(next)) engineRef.current.submitSoon()
        const manager = presenceRef.current
        if (manager && (tr.selectionSet || tr.docChanged)) {
          manager.publish({ from: next.selection.from, to: next.selection.to })
        }
        refresh(next)
      },
      handleKeyDown(_view, event) {
        if (event.key !== 'Enter' || event.shiftKey || event.metaKey || event.ctrlKey || event.altKey) return false
        const tr = view.state.tr.split(view.state.selection.from)
        if (!tr.docChanged) return false
        view.dispatch(tr)
        return true
      },
    })
    viewRef.current = view

    function refresh(state?: EditorState) {
      const manager = presenceRef.current
      const editorState = state ?? viewRef.current?.state
      const peer = manager?.peers()[0]
      setRemoteCaret(peer ? `${peer.name} ${caretLabel(peer.selection)}` : 'no remote caret')
      setMetrics({
        peers: manager?.peers().length ?? 0,
        applied: manager?.appliedSeq ?? 0,
        pending: manager && editorState ? sendableSteps(editorState)?.steps.length ?? 0 : 0,
      })
    }

    const startEditor = () => {
      if (disposed) return
      const transport = hub.connect<DocSelection>({ userId: profile.id, color: profile.color })
      const manager = new DocPresenceManager(transport, { path: SIM_ROOMS.docs, name: profile.name })
      presenceRef.current = manager
      manager.onChange(() => refresh())
      stopEditor = () => {
        manager.stop()
        transport.close()
        view.destroy()
        viewRef.current = null
        engineRef.current = null
      }
      void manager.start().then(async () => {
        if (disposed) return
        if (!manager.clientId) {
          setStatus('error')
          return
        }
        const engine = await manager.startSync({
          applyRemote(ops, remoteClientID) {
            const live = viewRef.current
            if (!live) throw new Error('editor missing')
            const steps = stepsFromOps(ops)
            const author = remoteClientID === manager.clientId ? clientID : remoteClientID
            live.dispatch(receiveTransaction(live.state, steps, steps.map(() => author)))
          },
          getPending() {
            const live = viewRef.current
            if (!live) return null
            const sendable = sendableSteps(live.state)
            if (!sendable || sendable.steps.length === 0) return null
            return { ops: opsFromSteps(sendable.steps) }
          },
          onAcked(ops) {
            const live = viewRef.current
            if (!live) return
            const steps = stepsFromOps(ops)
            live.dispatch(receiveTransaction(live.state, steps, steps.map(() => clientID)))
          },
          onResync() {
            setStatus('error')
          },
        })
        if (disposed) return
        if (!engine) {
          setStatus('error')
          return
        }
        engineRef.current = engine
        manager.publish({ from: view.state.selection.from, to: view.state.selection.to })
        setPresence(manager)
        refresh(view.state)
        setStatus('live')
      })
    }

    const startTimer = window.setTimeout(startEditor, profile.id === 'noah' ? 75 : 0)
    return () => {
      disposed = true
      window.clearTimeout(startTimer)
      if (stopEditor) stopEditor()
      else view.destroy()
    }
  }, [hub, profile])

  return (
    <SimEditorFrame
      profile={profile}
      status={status}
      presence={presence}
      metrics={metrics}
      extraFooter={<span>{remoteCaret}</span>}
    >
      <div ref={editorRef} className="collab-sim-editor__canvas collab-sim-editor__canvas--docs collab-prosemirror" data-testid={`collab-docs-editor-${profile.id}`} />
    </SimEditorFrame>
  )
}

export function CollabSimulatorDocs({ hub }: { hub: BrowserCollabHub }) {
  return (
    <>
      {SIM_EDITORS.map((profile) => <SimulatedDocsEditor key={profile.id} hub={hub} profile={profile} />)}
    </>
  )
}
