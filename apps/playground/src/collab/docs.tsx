import { useEffect, useRef, useState } from 'react'
import { collab, receiveTransaction, sendableSteps } from 'prosemirror-collab'
import { EditorState } from 'prosemirror-state'
import { EditorView } from 'prosemirror-view'
import {
  DocPresenceManager,
  type CollabTransport,
  type DocSelection,
  type DocSyncEngine,
  type PeerInfo,
} from '../../../../packages/collab/src/index.js'
import { createHttpCollabTransport, type HttpCollabTransport } from '../collabTransport'
import { initialCollabArtifact, useCollabComposition } from '../collabComposition'
import {
  COLLAB_API_BASE,
  CollabRoomChrome,
  defaultCollabName,
  fetchSampleWorkbook,
  mintCollabArtifact,
  type ConnectionState,
} from './roomChrome'
import { opsFromSteps, seedDocsCollabDoc, stepsFromOps } from './docsEditor'

const DOCS_PROOF = 'This tab mounts a live ProseMirror document and sends steps, ordered replay, reconnect catch-up, and caret presence {from,to} through injoffice-server over HTTP and server-sent events. It is not a native DOCX editor.'

function caretLabel(selection: DocSelection | null | undefined): string {
  if (!selection) return 'no caret'
  return `{from: ${selection.from}, to: ${selection.to}}`
}

export function CollabDocsPlaceholder() {
  const scope = useCollabComposition()
  const editorRef = useRef<HTMLDivElement | null>(null)
  const viewRef = useRef<EditorView | null>(null)
  const transportRef = useRef<HttpCollabTransport | null>(null)
  const managerRef = useRef<DocPresenceManager | null>(null)
  const engineRef = useRef<DocSyncEngine | null>(null)
  const collabClientIdRef = useRef(`docs-${crypto.randomUUID()}`)
  const [name, setName] = useState(defaultCollabName)
  const [artifact, setArtifact] = useState(() => initialCollabArtifact(scope, window.location.href))
  const [manager, setManager] = useState<DocPresenceManager | null>(null)
  const [connection, setConnection] = useState<ConnectionState>('idle')
  const [editorReady, setEditorReady] = useState(false)
  const [busy, setBusy] = useState(false)
  const [status, setStatus] = useState('Mint or paste an artifact ID, then join. The XLSX only identifies the room; this editor does not load a DOCX.')
  const [error, setError] = useState('')
  const [metrics, setMetrics] = useState({ peers: 0, applied: 0, pending: 0 })
  const [peers, setPeers] = useState<PeerInfo<DocSelection>[]>([])

  const refreshProtocol = (state?: EditorState) => {
    const nextManager = managerRef.current
    const editorState = state ?? viewRef.current?.state
    setPeers(nextManager?.peers() ?? [])
    setMetrics({
      peers: nextManager?.peers().length ?? 0,
      applied: nextManager?.appliedSeq ?? 0,
      pending: nextManager && editorState ? sendableSteps(editorState)?.steps.length ?? 0 : 0,
    })
  }

  useEffect(() => {
    const host = editorRef.current
    if (!host) return
    const clientID = collabClientIdRef.current
    const view = new EditorView(host, {
      state: EditorState.create({
        doc: seedDocsCollabDoc(),
        plugins: [collab({ version: 0, clientID })],
      }),
      dispatchTransaction(tr) {
        const next = view.state.apply(tr)
        view.updateState(next)
        const engine = engineRef.current
        if (engine && sendableSteps(next)) engine.submitSoon()
        const nextManager = managerRef.current
        if (nextManager && (tr.selectionSet || tr.docChanged)) {
          nextManager.publish({ from: next.selection.from, to: next.selection.to })
        }
        refreshProtocol(next)
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
    setEditorReady(true)
    return () => {
      managerRef.current?.stop()
      transportRef.current?.close()
      managerRef.current = null
      transportRef.current = null
      engineRef.current = null
      view.destroy()
      viewRef.current = null
    }
  }, [])

  const disconnect = () => {
    managerRef.current?.stop()
    transportRef.current?.close()
    managerRef.current = null
    transportRef.current = null
    engineRef.current = null
    setManager(null)
    setPeers([])
    setMetrics({ peers: 0, applied: 0, pending: 0 })
  }

  const joinRoom = async (path: string) => {
    const view = viewRef.current
    if (!view) {
      setError('The document editor is still starting. Try Join again in a moment.')
      return
    }
    setBusy(true)
    setConnection('joining')
    setError('')
    disconnect()
    try {
      const transport = await createHttpCollabTransport(COLLAB_API_BASE)
      transportRef.current = transport
      const nextManager = new DocPresenceManager(transport as unknown as CollabTransport<DocSelection>, { path, name })
      managerRef.current = nextManager
      nextManager.onChange(() => refreshProtocol())
      nextManager.onFileChanged(() => {
        setStatus('The backing artifact changed outside this editor. Remint the room before continuing.')
      })
      await nextManager.start()
      if (!nextManager.clientId) throw new Error('The sidecar did not join the collaboration room.')
      const clientID = collabClientIdRef.current
      const engine = await nextManager.startSync({
        applyRemote(ops, remoteClientID) {
          const live = viewRef.current
          if (!live) throw new Error('editor missing')
          const steps = stepsFromOps(ops)
          const author = remoteClientID === nextManager.clientId ? clientID : remoteClientID
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
          setConnection('error')
          setError('The operation log requires a document reload. Remint the artifact to restart this proof.')
        },
      })
      if (!engine) throw new Error('The sidecar did not start document sync.')
      engineRef.current = engine
      nextManager.publish({ from: view.state.selection.from, to: view.state.selection.to })
      setManager(nextManager)
      setArtifact(path)
      refreshProtocol(view.state)
      setConnection('live')
      setStatus(`Live as ${nextManager.state.me?.name ?? name}. Type in the document, then watch it appear in the second tab.`)
    } catch (cause) {
      disconnect()
      setConnection('error')
      setError(cause instanceof Error ? cause.message : String(cause))
    } finally {
      setBusy(false)
    }
  }

  const onMint = async (file: Blob) => {
    setBusy(true)
    setError('')
    try {
      const id = await mintCollabArtifact(file)
      setArtifact(id)
      await joinRoom(id)
    } catch (cause) {
      setConnection('error')
      setError(cause instanceof Error ? cause.message : String(cause))
      setBusy(false)
    }
  }

  return (
    <CollabRoomChrome
      title="Live ProseMirror protocol proof"
      connection={connection}
      name={name}
      onNameChange={setName}
      artifact={artifact}
      onArtifactChange={setArtifact}
      busy={busy}
      joinDisabled={!editorReady}
      onJoin={() => void joinRoom(artifact)}
      presence={manager}
      status={status}
      error={error}
      proof={DOCS_PROOF}
      editorTitle="Shared document"
      editorHint="Edit the ProseMirror surface directly. The uploaded XLSX identifies the room; this proof does not import DOCX contents."
      editorMeta={editorReady ? 'Editor ready' : 'Starting editor'}
      metrics={metrics}
      protocolExtra={(
        <>
          <h4>Peers</h4>
          {peers.length === 0 ? (
            <p className="native-muted">No other tabs in this room.</p>
          ) : (
            <ul className="collab-peer-list">
              {peers.map((peer) => (
                <li key={peer.client_id}>
                  <span>{peer.name}</span>
                  <code>{caretLabel(peer.selection)}</code>
                </li>
              ))}
            </ul>
          )}
        </>
      )}
      tryIt={(
        <ol>
          <li>Mint the sample workbook. It is only an opaque room key.</li>
          <li>Open the same URL in another tab.</li>
          <li>Type in the document and move the caret.</li>
          <li>Confirm the ordered text and remote {'{from,to}'} in both protocol panes.</li>
        </ol>
      )}
      actions={(
        <button
          type="button"
          className="workbench-button"
          disabled={busy || !editorReady}
          onClick={() => {
            void fetchSampleWorkbook()
              .then((blob) => onMint(blob))
              .catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)))
          }}
        >
          Mint sample
        </button>
      )}
    >
      <div ref={editorRef} className="collab-prosemirror" data-testid="collab-docs-editor" />
    </CollabRoomChrome>
  )
}

export default CollabDocsPlaceholder
