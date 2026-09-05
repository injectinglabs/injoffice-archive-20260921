import { useEffect, useRef, useState, type CSSProperties } from 'react'
import {
  DeckPresenceManager,
  type CollabTransport,
  type DeckSelection,
  type DeckSyncEngine,
} from '../../../../packages/collab/src/index.js'
import {
  BUILTIN_THEMES,
  DeckCanvasView,
  DeckEditorPanel,
  type DeckSpec,
} from '@injoffice/slides'
import { seedCollabDeck } from './slidesSeed'
import { createHttpCollabTransport, type HttpCollabTransport } from '../collabTransport'
import { parseCollabQuery } from '../collabScope'
import {
  COLLAB_API_BASE,
  CollabRoomChrome,
  defaultCollabName,
  fetchSampleWorkbook,
  mintCollabArtifact,
  type ConnectionState,
} from './roomChrome'

const SLIDES_PROOF = 'This tab mounts a live DeckSpec canvas and sends field updates, one-slide structure edits, ordered replay, reconnect catch-up, and slide presence through injoffice-server over HTTP and server-sent events.'

export function CollabSlidesPanel() {
  const transportRef = useRef<HttpCollabTransport | null>(null)
  const managerRef = useRef<DeckPresenceManager<DeckSpec> | null>(null)
  const engineRef = useRef<DeckSyncEngine<DeckSpec> | null>(null)
  const specRef = useRef<DeckSpec | null>(null)
  const [name, setName] = useState(defaultCollabName)
  const [artifact, setArtifact] = useState(() => parseCollabQuery(window.location.href).artifact)
  const [manager, setManager] = useState<DeckPresenceManager<DeckSpec> | null>(null)
  const [connection, setConnection] = useState<ConnectionState>('idle')
  const [busy, setBusy] = useState(false)
  const [status, setStatus] = useState('Mint or paste an artifact ID, then join. The XLSX only identifies the room; the deck is seeded locally.')
  const [error, setError] = useState('')
  const [metrics, setMetrics] = useState({ peers: 0, applied: 0, pending: 0 })
  const [spec, setSpec] = useState(seedCollabDeck)
  const [at, setAt] = useState(0)
  const [shapeKey, setShapeKey] = useState<string | null>(null)
  specRef.current = spec

  useEffect(() => () => {
    managerRef.current?.stop()
    transportRef.current?.close()
    managerRef.current = null
    transportRef.current = null
    engineRef.current = null
  }, [])

  useEffect(() => {
    if (at >= spec.slides.length) setAt(Math.max(0, spec.slides.length - 1))
  }, [spec, at])

  useEffect(() => {
    const slide = spec.slides[at]
    if (manager && slide) manager.publish({ slideId: slide.id, shapeKey })
  }, [manager, spec, at, shapeKey])

  const disconnect = () => {
    managerRef.current?.stop()
    transportRef.current?.close()
    managerRef.current = null
    transportRef.current = null
    engineRef.current = null
    setManager(null)
    setMetrics({ peers: 0, applied: 0, pending: 0 })
  }

  const applyDeck = (next: DeckSpec) => {
    const engine = engineRef.current
    if (!engine) {
      setSpec(next)
      return
    }
    if (!engine.change(next)) {
      setStatus('That edit rebuilt the outline and was not synced. Use a one-slide add, delete, or move.')
    }
  }

  const joinRoom = async (path: string) => {
    setBusy(true)
    setConnection('joining')
    setError('')
    disconnect()
    try {
      const transport = await createHttpCollabTransport(COLLAB_API_BASE)
      transportRef.current = transport
      const nextManager = new DeckPresenceManager<DeckSpec>(transport as unknown as CollabTransport<DeckSelection>, { path, name })
      managerRef.current = nextManager
      nextManager.onChange(() => {
        setMetrics({
          peers: nextManager.peers().length,
          applied: nextManager.appliedSeq,
          pending: 0,
        })
      })
      nextManager.onFileChanged(() => {
        setStatus('The backing artifact changed outside this editor. Remint the room before continuing.')
      })
      await nextManager.start()
      if (!nextManager.clientId) throw new Error('The sidecar did not join the collaboration room.')
      const engine = await nextManager.startSync(specRef.current ?? spec, {
        onDeck: setSpec,
        onResync: () => {
          setConnection('error')
          setError('The operation log requires a deck reload. Remint the artifact to restart this proof.')
        },
      })
      if (!engine) throw new Error('The sidecar did not start DeckSpec sync.')
      engineRef.current = engine
      setManager(nextManager)
      setArtifact(path)
      setMetrics({ peers: nextManager.peers().length, applied: nextManager.appliedSeq, pending: 0 })
      setConnection('live')
      setStatus(`Live as ${nextManager.state.me?.name ?? name}. Edit a slide, then watch it appear in the second tab.`)
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

  const theme = BUILTIN_THEMES.find((item) => item.id === spec.theme) ?? BUILTIN_THEMES[0]!

  return (
    <CollabRoomChrome
      title="Live DeckSpec protocol proof"
      connection={connection}
      name={name}
      onNameChange={setName}
      artifact={artifact}
      onArtifactChange={setArtifact}
      busy={busy}
      onJoin={() => void joinRoom(artifact)}
      presence={manager}
      status={status}
      error={error}
      proof={SLIDES_PROOF}
      editorTitle="Shared deck"
      editorHint="Edit the DeckSpec directly. The uploaded XLSX identifies the room; this proof does not import its contents."
      editorMeta={`${spec.slides.length} ${spec.slides.length === 1 ? 'slide' : 'slides'}`}
      metrics={metrics}
      tryIt={(
        <ol>
          <li>Mint the sample workbook. It is only an opaque room key.</li>
          <li>Open the same URL in another tab.</li>
          <li>Edit a title or drag a shape.</li>
          <li>Confirm the ordered edit and remote slide presence in both tabs.</li>
        </ol>
      )}
      actions={(
        <button
          type="button"
          className="workbench-button"
          disabled={busy}
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
      <div className="split" style={{ flex: 1, minHeight: 500 }}>
        <div
          className="split-main slides-workspace"
          style={{ '--slides-workspace-background': theme.background } as CSSProperties}
        >
          <DeckCanvasView
            spec={spec}
            at={at}
            onAtChange={setAt}
            width={720}
            editable
            onShapeOverride={(slideIndex, key, patch) => {
              const slide = spec.slides[slideIndex]
              if (!slide) return
              setShapeKey(key)
              applyDeck({
                ...spec,
                slides: spec.slides.map((item, index) =>
                  index === slideIndex
                    ? { ...item, shapeOverrides: { ...item.shapeOverrides, [key]: { ...item.shapeOverrides?.[key], ...patch } } }
                    : item,
                ),
              })
            }}
          />
        </div>
        <aside className="split-side">
          <DeckEditorPanel spec={spec} at={at} onChange={applyDeck} onSelect={setAt} />
        </aside>
      </div>
    </CollabRoomChrome>
  )
}

export default CollabSlidesPanel
