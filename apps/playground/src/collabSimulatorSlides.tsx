import { useEffect, useRef, useState, type CSSProperties } from 'react'
import {
  DeckPresenceManager,
  type DeckSelection,
  type DeckSyncEngine,
} from '../../../packages/collab/src/index.js'
import { BUILTIN_THEMES, DeckCanvasView, type DeckSpec } from '@injoffice/slides'
import type { BrowserCollabHub } from './browserCollabTransport'
import { seedCollabDeck } from './collab/slidesSeed'
import { SIM_EDITORS, SIM_ROOMS, SimEditorFrame, type SimEditorProfile, type SimMetrics } from './collabSimChrome'

function cloneDeck(spec: DeckSpec): DeckSpec {
  return JSON.parse(JSON.stringify(spec)) as DeckSpec
}

function SimulatedSlidesEditor({
  hub,
  profile,
  seed,
}: {
  hub: BrowserCollabHub
  profile: SimEditorProfile
  seed: DeckSpec
}) {
  const specRef = useRef<DeckSpec>(cloneDeck(seed))
  const engineRef = useRef<DeckSyncEngine<DeckSpec> | null>(null)
  const [spec, setSpec] = useState(() => cloneDeck(seed))
  const [at, setAt] = useState(0)
  const [shapeKey, setShapeKey] = useState<string | null>(null)
  const [status, setStatus] = useState<'starting' | 'live' | 'error'>('starting')
  const [metrics, setMetrics] = useState<SimMetrics>({ peers: 0, applied: 0, pending: 0 })
  const [presence, setPresence] = useState<DeckPresenceManager<DeckSpec> | null>(null)
  const [hint, setHint] = useState('Connecting…')
  specRef.current = spec

  useEffect(() => {
    if (at >= spec.slides.length) setAt(Math.max(0, spec.slides.length - 1))
  }, [spec, at])

  useEffect(() => {
    const slide = spec.slides[at]
    if (presence && slide) presence.publish({ slideId: slide.id, shapeKey })
  }, [presence, spec, at, shapeKey])

  useEffect(() => {
    let disposed = false
    let stopEditor: (() => void) | null = null
    const startEditor = () => {
      if (disposed) return
      const transport = hub.connect<DeckSelection>({ userId: profile.id, color: profile.color })
      const manager = new DeckPresenceManager<DeckSpec>(transport, { path: SIM_ROOMS.slides, name: profile.name })
      const refresh = () => setMetrics({
        peers: manager.peers().length,
        applied: manager.appliedSeq,
        pending: 0,
      })
      manager.onChange(refresh)
      stopEditor = () => {
        manager.stop()
        transport.close()
        engineRef.current = null
      }
      void manager.start().then(async () => {
        if (disposed) return
        if (!manager.clientId) {
          setStatus('error')
          return
        }
        const engine = await manager.startSync(specRef.current, {
          onDeck: setSpec,
          onResync: () => setStatus('error'),
        })
        if (disposed) return
        if (!engine) {
          setStatus('error')
          return
        }
        engineRef.current = engine
        setPresence(manager)
        refresh()
        setStatus('live')
        setHint('Edit the title or a shape.')
      })
    }
    const startTimer = window.setTimeout(startEditor, profile.id === 'noah' ? 75 : 0)
    return () => {
      disposed = true
      window.clearTimeout(startTimer)
      stopEditor?.()
    }
  }, [hub, profile])

  const applyDeck = (next: DeckSpec) => {
    const engine = engineRef.current
    if (!engine) {
      setSpec(next)
      return
    }
    if (!engine.change(next)) setHint('That edit rebuilt the outline and was not synced. Change the title or one shape.')
    else setHint('Edit queued for the ordered log.')
  }

  const theme = BUILTIN_THEMES.find((item) => item.id === spec.theme) ?? BUILTIN_THEMES[0]!
  const peer = presence?.peers()[0]
  const peerSlide = peer?.selection?.slideId ? `${peer.name} on ${peer.selection.slideId}` : 'no remote slide'

  return (
    <SimEditorFrame
      profile={profile}
      status={status}
      presence={presence}
      metrics={metrics}
      extraFooter={<span>{peerSlide}</span>}
    >
      <div className="collab-sim-editor__canvas collab-sim-editor__canvas--slides" data-testid={`collab-slides-editor-${profile.id}`}>
        <div className="collab-sim-slides-tools">
          <label>
            Title
            <input
              value={spec.title}
              onChange={(event) => applyDeck({ ...spec, title: event.target.value })}
              aria-label={`${profile.name} deck title`}
            />
          </label>
          <button type="button" className="workbench-button" disabled={at <= 0} onClick={() => setAt((current) => Math.max(0, current - 1))}>Previous</button>
          <button type="button" className="workbench-button" disabled={at >= spec.slides.length - 1} onClick={() => setAt((current) => Math.min(spec.slides.length - 1, current + 1))}>Next</button>
          <span>{at + 1} / {spec.slides.length}</span>
        </div>
        <div className="collab-sim-slides-canvas" style={{ '--slides-workspace-background': theme.background } as CSSProperties}>
          <DeckCanvasView
            spec={spec}
            at={at}
            onAtChange={setAt}
            width={420}
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
        <p className="collab-sim-slides-hint">{hint}</p>
      </div>
    </SimEditorFrame>
  )
}

export function CollabSimulatorSlides({ hub }: { hub: BrowserCollabHub }) {
  const [seed] = useState(seedCollabDeck)
  return (
    <>
      {SIM_EDITORS.map((profile) => <SimulatedSlidesEditor key={profile.id} hub={hub} profile={profile} seed={seed} />)}
    </>
  )
}
