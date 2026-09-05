import { useEffect, useRef, useState } from 'react'
import {
  PdfPresenceManager,
  type DocSyncEngine,
  type PdfSelection,
} from '../../../packages/collab/src/index.js'
import { encodePdfCollabOperation, type PdfCollabOperation } from '../../../packages/pdf/src/collab'
import { PdfViewerDocument, renderPageToCanvas } from '../../../packages/pdf/src/viewer'
import type { BrowserCollabHub } from './browserCollabTransport'
import {
  applyPdfCollabOps,
  inspectCollabPdf,
  makeCollabPdfSample,
  SAMPLE_HIGHLIGHT_QUADS,
  type CollabPdfAnnot,
} from './collab/pdf'
import { SIM_EDITORS, SIM_ROOMS, SimEditorFrame, type SimEditorProfile, type SimMetrics } from './collabSimChrome'

const MEMO_FIELD = 'shared.memo'
const AGREE_FIELD = 'shared.agree'

function clampPdfPage(page: number, pageCount: number): number {
  if (pageCount < 1) return 1
  return Math.max(1, Math.min(Math.trunc(page), pageCount))
}

function SimulatedPdfEditor({
  hub,
  profile,
  seed,
}: {
  hub: BrowserCollabHub
  profile: SimEditorProfile
  seed: Uint8Array
}) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const engineRef = useRef<DocSyncEngine | null>(null)
  const viewerRef = useRef<PdfViewerDocument | null>(null)
  const bytesRef = useRef<Uint8Array>(new Uint8Array(seed))
  const pendingRef = useRef<unknown[]>([])
  const applyChainRef = useRef(Promise.resolve())
  const renderQueueRef = useRef(Promise.resolve())
  const renderGenerationRef = useRef(0)
  const [bytes, setBytes] = useState(() => new Uint8Array(seed))
  const [viewer, setViewer] = useState<PdfViewerDocument | null>(null)
  const [page, setPage] = useState(1)
  const [pageCount, setPageCount] = useState(0)
  const [annots, setAnnots] = useState<CollabPdfAnnot[]>([])
  const [memo, setMemo] = useState('')
  const [agree, setAgree] = useState(false)
  const [noteDraft, setNoteDraft] = useState('')
  const [status, setStatus] = useState<'starting' | 'live' | 'error'>('starting')
  const [metrics, setMetrics] = useState<SimMetrics>({ peers: 0, applied: 0, pending: 0 })
  const [presence, setPresence] = useState<PdfPresenceManager | null>(null)
  bytesRef.current = bytes

  const runExclusive = (work: () => Promise<void>): Promise<void> => {
    const next = applyChainRef.current.then(work, work)
    applyChainRef.current = next.then(() => undefined, () => undefined)
    return next
  }

  const refreshMetrics = (manager: PdfPresenceManager | null) => {
    setMetrics({
      peers: manager?.peers().length ?? 0,
      applied: manager?.appliedSeq ?? 0,
      pending: pendingRef.current.length,
    })
  }

  useEffect(() => {
    let cancelled = false
    void inspectCollabPdf(bytes).then((info) => {
      if (cancelled) return
      setAnnots(info.annots)
      const memoField = info.forms.find((field) => field.name === MEMO_FIELD)
      const agreeField = info.forms.find((field) => field.name === AGREE_FIELD)
      if (memoField?.value !== undefined) setMemo(memoField.value)
      if (agreeField?.checked !== undefined) setAgree(agreeField.checked)
      const note = info.annots.find((annot) => annot.subtype === 'note')
      if (note?.contents !== undefined) setNoteDraft(note.contents)
    }).catch(() => {
      if (!cancelled) setAnnots([])
    })
    return () => {
      cancelled = true
    }
  }, [bytes])

  useEffect(() => {
    let cancelled = false
    let ownedViewer: PdfViewerDocument | null = null
    renderGenerationRef.current += 1
    viewerRef.current = null
    setViewer(null)
    setPageCount(0)
    void PdfViewerDocument.load(bytes).then((loaded) => {
      ownedViewer = loaded
      if (cancelled) return
      viewerRef.current = loaded
      setPageCount(loaded.pageCount)
      setPage((current) => clampPdfPage(current, loaded.pageCount))
      setViewer(loaded)
    }).catch(() => {
      if (!cancelled) setStatus('error')
    })
    return () => {
      cancelled = true
      renderGenerationRef.current += 1
      if (ownedViewer) {
        if (viewerRef.current === ownedViewer) viewerRef.current = null
        void ownedViewer.destroy().catch(() => undefined)
      }
    }
  }, [bytes])

  useEffect(() => {
    if (!viewer || pageCount < 1) return
    let cancelled = false
    renderGenerationRef.current += 1
    const selectedPage = clampPdfPage(page, pageCount)
    const render = renderQueueRef.current
      .catch(() => undefined)
      .then(async () => {
        const canvas = canvasRef.current
        if (!canvas || cancelled || viewerRef.current !== viewer) return
        await renderPageToCanvas(viewer, selectedPage, canvas, 1)
      })
    renderQueueRef.current = render
    return () => {
      cancelled = true
    }
  }, [page, pageCount, viewer])

  useEffect(() => {
    if (!presence) return
    presence.publish({ page, annotLocalId: null })
  }, [presence, page])

  useEffect(() => {
    let disposed = false
    let stopEditor: (() => void) | null = null
    const startEditor = () => {
      if (disposed) return
      const transport = hub.connect<PdfSelection>({ userId: profile.id, color: profile.color })
      const manager = new PdfPresenceManager(transport, { path: SIM_ROOMS.pdf, name: profile.name })
      manager.onChange(() => refreshMetrics(manager))
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
        const engine = await manager.startSync({
          applyRemote(ops) {
            return runExclusive(async () => {
              const next = await applyPdfCollabOps(bytesRef.current, ops)
              bytesRef.current = next
              setBytes(Uint8Array.from(next))
              refreshMetrics(manager)
            })
          },
          getPending() {
            const ops = pendingRef.current
            return ops.length > 0 ? { ops: ops.slice() } : null
          },
          onAcked(ops) {
            pendingRef.current = pendingRef.current.slice(ops.length)
            refreshMetrics(manager)
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
        setPresence(manager)
        refreshMetrics(manager)
        setStatus('live')
      })
    }
    const startTimer = window.setTimeout(startEditor, profile.id === 'noah' ? 75 : 0)
    return () => {
      disposed = true
      window.clearTimeout(startTimer)
      stopEditor?.()
    }
  }, [hub, profile])

  const submitLocal = async (operation: PdfCollabOperation) => {
    if (!engineRef.current) return
    const encoded = encodePdfCollabOperation(operation)
    await runExclusive(async () => {
      const next = await applyPdfCollabOps(bytesRef.current, [encoded])
      bytesRef.current = next
      setBytes(Uint8Array.from(next))
      pendingRef.current = [...pendingRef.current, encoded]
      engineRef.current?.submitSoon()
      refreshMetrics(presence)
    })
  }

  const note = annots.find((annot) => annot.subtype === 'note')
  const peer = presence?.peers()[0]
  const peerPage = peer?.selection?.page ? `${peer.name} on page ${peer.selection.page}` : 'no remote page'
  const ready = viewer !== null && pageCount > 0 && status === 'live'

  return (
    <SimEditorFrame
      profile={profile}
      status={status}
      presence={presence}
      metrics={metrics}
      extraFooter={<span>{peerPage}</span>}
    >
      <div className="collab-sim-editor__canvas collab-sim-editor__canvas--pdf" data-testid={`collab-pdf-editor-${profile.id}`}>
        <div className="collab-sim-pdf-tools">
          <button type="button" className="workbench-button" disabled={!ready || page <= 1} onClick={() => setPage((current) => clampPdfPage(current - 1, pageCount))}>Previous</button>
          <span>Page {page} / {pageCount || '—'}</span>
          <button type="button" className="workbench-button" disabled={!ready || page >= pageCount} onClick={() => setPage((current) => clampPdfPage(current + 1, pageCount))}>Next</button>
          <button
            type="button"
            className="workbench-button workbench-button--primary"
            disabled={!ready}
            onClick={() => void submitLocal({
              kind: 'annotation.markup',
              value: { page, type: 'highlight', color: [1, 0.92, 0.2], quads: SAMPLE_HIGHLIGHT_QUADS },
            })}
          >
            Highlight
          </button>
        </div>
        <canvas ref={canvasRef} className="pdf-canvas collab-sim-pdf-canvas" aria-label={`${profile.name} PDF page ${page}`} />
        <div className="collab-sim-pdf-tools">
          <label>
            Sticky note
            <input
              value={noteDraft}
              disabled={!ready || !note}
              onChange={(event) => setNoteDraft(event.target.value)}
              onBlur={() => {
                if (!note || noteDraft === (note.contents ?? '')) return
                void submitLocal({
                  kind: 'annotation.note',
                  value: {
                    page: note.page,
                    objNum: note.objNum,
                    rect: note.rect,
                    oldContents: note.contents ?? '',
                    contents: noteDraft,
                  },
                })
              }}
            />
          </label>
          <label>
            Memo
            <input
              value={memo}
              disabled={!ready}
              onChange={(event) => setMemo(event.target.value)}
              onBlur={() => {
                void submitLocal({ kind: 'form.value', value: { name: MEMO_FIELD, kind: 'text', value: memo } })
              }}
            />
          </label>
          <label>
            <input
              type="checkbox"
              checked={agree}
              disabled={!ready}
              onChange={(event) => {
                const checked = event.target.checked
                setAgree(checked)
                void submitLocal({ kind: 'form.value', value: { name: AGREE_FIELD, kind: 'checkbox', checked } })
              }}
            />
            Agree
          </label>
        </div>
      </div>
    </SimEditorFrame>
  )
}

export function CollabSimulatorPdf({ hub }: { hub: BrowserCollabHub }) {
  const [seed, setSeed] = useState<Uint8Array | null>(null)
  const [error, setError] = useState('')

  useEffect(() => {
    let cancelled = false
    void makeCollabPdfSample().then((bytes) => {
      if (!cancelled) setSeed(bytes)
    }).catch((cause: unknown) => {
      if (!cancelled) setError(cause instanceof Error ? cause.message : String(cause))
    })
    return () => {
      cancelled = true
    }
  }, [])

  if (error) return <p className="native-error workbench-callout workbench-callout--error" role="alert">{error}</p>
  if (!seed) return <p className="collab-sim-empty">Loading the shared PDF…</p>
  return (
    <>
      {SIM_EDITORS.map((profile) => <SimulatedPdfEditor key={profile.id} hub={hub} profile={profile} seed={seed} />)}
    </>
  )
}
