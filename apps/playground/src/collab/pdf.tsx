import { useEffect, useRef, useState } from 'react'
import {
  PDFArray,
  PDFCheckBox,
  PDFDict,
  PDFDocument,
  PDFDropdown,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFOptionList,
  PDFRadioGroup,
  PDFRef,
  PDFString,
  PDFTextField,
  StandardFonts,
  rgb,
} from 'pdf-lib'
import pdfWorkerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import {
  PdfPresenceManager,
  PresenceStack,
  type CollabTransport,
  type DocSyncEngine,
  type PdfSelection,
  type PeerInfo,
} from '../../../../packages/collab/src/index.js'
// Import the browser-safe annotation leaves directly. The annotate barrel also
// exports destructive form redaction, whose PDFium loader intentionally uses
// Node filesystem APIs and therefore cannot be evaluated by a browser bundle.
import { applyAnnotDeletes } from '../../../../packages/pdf/src/annotate/annotDelete'
import { applyDrawings, applyNoteEdits } from '../../../../packages/pdf/src/annotate/drawing'
import { applyFormValues } from '../../../../packages/pdf/src/annotate/forms'
import { applyMarkups } from '../../../../packages/pdf/src/annotate/markup'
import {
  decodePdfCollabOperation,
  encodePdfCollabOperation,
  type PdfCollabOperation,
} from '../../../../packages/pdf/src/collab'
import {
  configurePdfWorker,
  PdfViewerDocument,
  renderPageToCanvas,
} from '../../../../packages/pdf/src/viewer'
import { COLLAB_COPY, parseCollabQuery, writeCollabQuery } from '../collabScope'
import { createHttpCollabTransport, type HttpCollabTransport } from '../collabTransport'

configurePdfWorker(pdfWorkerUrl)

type ConnectionState = 'idle' | 'joining' | 'live' | 'error'

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const SAMPLE_XLSX = `${import.meta.env.BASE_URL}collab-demo.xlsx`
const COLLAB_API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').replace(/\/$/, '')
const SAMPLE_NOTE = 'Shared note — edit me'
const MEMO_FIELD = 'shared.memo'
const AGREE_FIELD = 'shared.agree'
export const SAMPLE_HIGHLIGHT_QUADS: number[][] = [[72, 654, 360, 654, 72, 636, 360, 636]]

const PDF_PROOF =
  'This tab mounts a live PDF annotator and sends highlight, note, form-value, and annotation-delete operations, ordered replay, reconnect catch-up, and page presence through injoffice-server over HTTP and server-sent events. Annotation and form values are collaborated; page, text, and image edits are not.'

const ANNOT_SUBTYPES = {
  Highlight: 'highlight',
  Underline: 'underline',
  StrikeOut: 'strikeout',
  Text: 'note',
} as const

export type CollabPdfAnnot = {
  page: number
  objNum: number
  subtype: 'highlight' | 'underline' | 'strikeout' | 'note'
  rect: [number, number, number, number]
  contents?: string
}

export type CollabPdfFormField = {
  name: string
  kind: 'text' | 'checkbox' | 'radio' | 'choice'
  value?: string
  checked?: boolean
}

function defaultCollabName(): string {
  const stored = sessionStorage.getItem('injoffice-collab-name')
  if (stored) return stored
  return `Tab ${Math.floor(Math.random() * 90 + 10)}`
}



function connectionLabel(connection: ConnectionState): string {
  return connection === 'live' ? 'Connected' : connection === 'joining' ? 'Joining' : connection === 'error' ? 'Not connected' : 'Local only'
}

function clampPdfPage(page: number, pageCount: number): number {
  if (pageCount < 1) return 1
  return Math.max(1, Math.min(Math.trunc(page), pageCount))
}

function errorMessage(error: unknown): string {
  return error instanceof Error && error.message ? error.message : String(error)
}

function sidecarHint(action: string): Error {
  return new Error(`${action} failed. Start the local sidecar with \`go run ./cmd/injoffice-server\` from go/injoffice-server.`)
}

async function artifactIdFrom(response: Response, action: string): Promise<string> {
  const detail = await response.text()
  if (!response.ok) throw new Error(`${action} failed (${response.status}): ${detail}`)
  const id = response.headers.get('X-InjOffice-Artifact-Id')?.trim()
  if (!id) throw new Error(`${action} did not return an opaque artifact ID.`)
  return id
}

async function mintViaXlsxExtract(): Promise<string> {
  const sample = await fetch(SAMPLE_XLSX)
  if (!sample.ok) throw new Error(`Sample workbook missing (${sample.status}).`)
  const file = await sample.blob()
  let response: Response
  try {
    response = await fetch(`${COLLAB_API_BASE}/v1/xlsx/extract`, {
      method: 'POST',
      headers: { 'Content-Type': XLSX_TYPE },
      body: file,
    })
  } catch {
    throw sidecarHint('Extract')
  }
  return artifactIdFrom(response, 'Extract')
}

export async function mintCollabPdfArtifact(bytes: Uint8Array): Promise<string> {
  const body = Uint8Array.from(bytes)
  let response: Response
  try {
    response = await fetch(`${COLLAB_API_BASE}/v1/artifacts`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/pdf' },
      body,
    })
  } catch {
    throw sidecarHint('Mint')
  }
  if (response.status === 404) return mintViaXlsxExtract()
  return artifactIdFrom(response, 'Mint')
}

export async function makeCollabPdfSample(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.TimesRoman)
  const bold = await doc.embedFont(StandardFonts.HelveticaBold)
  const page = doc.addPage([612, 792])
  page.drawText('Shared PDF annotation proof', { x: 72, y: 720, size: 24, font: bold, color: rgb(0.1, 0.13, 0.22) })
  page.drawText('Highlight this shared sentence.', { x: 72, y: 640, size: 14, font: bold, color: rgb(0.2, 0.24, 0.31) })
  page.drawText('Edit the sticky note, fill the form, or delete a markup.', { x: 72, y: 610, size: 11, font, color: rgb(0.2, 0.24, 0.31) })
  page.drawText('Page, text, and image edits are not collaborated.', { x: 72, y: 592, size: 11, font, color: rgb(0.2, 0.24, 0.31) })
  page.drawText('Shared memo', { x: 72, y: 548, size: 10, font })
  const form = doc.getForm()
  form.createTextField(MEMO_FIELD).addToPage(page, { x: 72, y: 520, width: 280, height: 22 })
  form.createCheckBox(AGREE_FIELD).addToPage(page, { x: 72, y: 488, width: 16, height: 16 })
  page.drawText('Agree', { x: 94, y: 490, size: 11, font })
  page.drawText('InjOffice PDF annotation collab sample · page 1', { x: 72, y: 54, size: 9, font })

  const page2 = doc.addPage([612, 792])
  page2.drawText('Page presence', { x: 72, y: 720, size: 24, font: bold, color: rgb(0.1, 0.13, 0.22) })
  page2.drawText('Change pages to show peers which page you are on.', { x: 72, y: 668, size: 12, font })
  page2.drawText('InjOffice PDF annotation collab sample · page 2', { x: 72, y: 54, size: 9, font })

  return applyDrawings(await doc.save(), [{
    kind: 'note',
    page: 1,
    color: [1, 0.85, 0.2],
    at: [500, 700],
    contents: SAMPLE_NOTE,
    author: 'InjOffice',
  }])
}

export async function applyPdfCollabOps(bytes: Uint8Array, ops: unknown[]): Promise<Uint8Array> {
  let current = bytes
  for (const raw of ops) {
    const decoded = decodePdfCollabOperation(raw)
    if (!decoded.ok) continue
    const { operation } = decoded
    if (operation.kind === 'annotation.markup') current = await applyMarkups(current, [operation.value])
    else if (operation.kind === 'annotation.note') current = await applyNoteEdits(current, [operation.value])
    else if (operation.kind === 'annotation.delete') current = await applyAnnotDeletes(current, [operation.value])
    else current = (await applyFormValues(current, [operation.value])).bytes
  }
  return current
}

function rectOf(dict: PDFDict): [number, number, number, number] | null {
  const rect = dict.lookupMaybe(PDFName.of('Rect'), PDFArray)
  if (!rect || rect.size() !== 4) return null
  const values = [0, 1, 2, 3].map((index) => rect.lookupMaybe(index, PDFNumber)?.asNumber() ?? NaN)
  if (values.some((value) => !Number.isFinite(value))) return null
  return values as [number, number, number, number]
}

function contentsOf(dict: PDFDict): string | undefined {
  const value = dict.lookup(PDFName.of('Contents'))
  return value instanceof PDFString || value instanceof PDFHexString ? value.decodeText() : undefined
}

export async function inspectCollabPdf(bytes: Uint8Array): Promise<{ annots: CollabPdfAnnot[]; forms: CollabPdfFormField[] }> {
  const doc = await PDFDocument.load(bytes)
  const annots: CollabPdfAnnot[] = []
  doc.getPages().forEach((page, pageIndex) => {
    const arr = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!arr) return
    for (let index = 0; index < arr.size(); index += 1) {
      const ref = arr.get(index)
      if (!(ref instanceof PDFRef)) continue
      const dict = doc.context.lookupMaybe(ref, PDFDict)
      if (!dict) continue
      const subtypeName = dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()
      const subtype = subtypeName ? ANNOT_SUBTYPES[subtypeName as keyof typeof ANNOT_SUBTYPES] : undefined
      if (!subtype) continue
      const rect = rectOf(dict)
      if (!rect) continue
      annots.push({ page: pageIndex + 1, objNum: ref.objectNumber, subtype, rect, contents: contentsOf(dict) })
    }
  })

  const forms: CollabPdfFormField[] = []
  for (const field of doc.getForm().getFields()) {
    const name = field.getName()
    if (field instanceof PDFTextField) forms.push({ name, kind: 'text', value: field.getText() ?? '' })
    else if (field instanceof PDFCheckBox) forms.push({ name, kind: 'checkbox', checked: field.isChecked() })
    else if (field instanceof PDFRadioGroup) forms.push({ name, kind: 'radio', value: field.getSelected() ?? '' })
    else if (field instanceof PDFDropdown || field instanceof PDFOptionList) forms.push({ name, kind: 'choice', value: field.getSelected()?.[0] ?? '' })
  }
  return { annots, forms }
}

export function CollabPdfPanel() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const transportRef = useRef<HttpCollabTransport | null>(null)
  const managerRef = useRef<PdfPresenceManager | null>(null)
  const engineRef = useRef<DocSyncEngine | null>(null)
  const viewerRef = useRef<PdfViewerDocument | null>(null)
  const renderQueueRef = useRef(Promise.resolve())
  const renderGenerationRef = useRef(0)
  const bytesRef = useRef<Uint8Array | null>(null)
  const seedRef = useRef<Uint8Array | null>(null)
  const pendingRef = useRef<unknown[]>([])
  const applyChainRef = useRef(Promise.resolve())
  const [name, setName] = useState(defaultCollabName)
  const [artifact, setArtifact] = useState(() => parseCollabQuery(window.location.href).artifact)
  const memoDirtyRef = useRef(false)
  const agreeDirtyRef = useRef(false)
  const noteDirtyRef = useRef(false)
  const [manager, setManager] = useState<PdfPresenceManager | null>(null)
  const [peers, setPeers] = useState<PeerInfo<PdfSelection>[]>([])
  const [connection, setConnection] = useState<ConnectionState>('idle')
  const [busy, setBusy] = useState(false)
  const [status, setStatus] = useState('Mint or paste an artifact ID, then join. The PDF is seeded locally; the artifact is only an opaque room key.')
  const [error, setError] = useState('')
  const [metrics, setMetrics] = useState({ peers: 0, applied: 0, pending: 0 })
  const [bytes, setBytes] = useState<Uint8Array | null>(null)
  const [viewer, setViewer] = useState<PdfViewerDocument | null>(null)
  const [page, setPage] = useState(1)
  const [pageCount, setPageCount] = useState(0)
  const [annots, setAnnots] = useState<CollabPdfAnnot[]>([])
  const [selectedAnnot, setSelectedAnnot] = useState<CollabPdfAnnot | null>(null)
  const [noteDraft, setNoteDraft] = useState(SAMPLE_NOTE)
  const [memo, setMemo] = useState('')
  const [agree, setAgree] = useState(false)
  bytesRef.current = bytes

  const runExclusive = (work: () => Promise<void>): Promise<void> => {
    const next = applyChainRef.current.then(work, work)
    applyChainRef.current = next.then(() => undefined, () => undefined)
    return next
  }

  const refreshMetrics = (nextManager: PdfPresenceManager | null = managerRef.current) => {
    setMetrics({
      peers: nextManager?.peers().length ?? 0,
      applied: nextManager?.appliedSeq ?? 0,
      pending: pendingRef.current.length,
    })
    setPeers(nextManager?.peers() ?? [])
  }

  useEffect(() => () => {
    managerRef.current?.stop()
    transportRef.current?.close()
    managerRef.current = null
    transportRef.current = null
    engineRef.current = null
    void viewerRef.current?.destroy().catch(() => undefined)
  }, [])

  useEffect(() => {
    sessionStorage.setItem('injoffice-collab-name', name)
  }, [name])

  useEffect(() => {
    writeCollabQuery({ artifact, format: 'pdf' })
  }, [artifact])

  useEffect(() => {
    let cancelled = false
    void makeCollabPdfSample()
      .then((sample) => {
        if (cancelled) return
        seedRef.current = sample
        bytesRef.current = sample
        setBytes(sample)
      })
      .catch((reason: unknown) => {
        if (!cancelled) setError(`Could not generate the sample PDF: ${errorMessage(reason)}`)
      })
    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (!bytes) return
    let cancelled = false
    void inspectCollabPdf(bytes)
      .then((info) => {
        if (cancelled) return
        setAnnots(info.annots)
        const memoField = info.forms.find((field) => field.name === MEMO_FIELD)
        const agreeField = info.forms.find((field) => field.name === AGREE_FIELD)
        if (memoField?.value !== undefined && !memoDirtyRef.current) setMemo(memoField.value)
        if (agreeField?.checked !== undefined && !agreeDirtyRef.current) setAgree(agreeField.checked)
        const note = info.annots.find((annot) => annot.subtype === 'note')
        if (note?.contents !== undefined && !noteDirtyRef.current) setNoteDraft(note.contents)
        setSelectedAnnot((current) => {
          if (!current) return null
          return info.annots.find((annot) => annot.page === current.page && annot.objNum === current.objNum) ?? null
        })
      })
      .catch(() => {
        if (!cancelled) setAnnots([])
      })
    return () => {
      cancelled = true
    }
  }, [bytes])

  useEffect(() => {
    if (!bytes) return
    let cancelled = false
    let ownedViewer: PdfViewerDocument | null = null
    renderGenerationRef.current += 1
    viewerRef.current = null
    setViewer(null)
    setPageCount(0)
    const canvas = canvasRef.current
    if (canvas) {
      canvas.width = 0
      canvas.height = 0
      canvas.style.width = '0px'
      canvas.style.height = '0px'
    }

    void (async () => {
      try {
        const loaded = await PdfViewerDocument.load(bytes)
        ownedViewer = loaded
        if (cancelled) return
        viewerRef.current = loaded
        setPageCount(loaded.pageCount)
        setPage((current) => clampPdfPage(current, loaded.pageCount))
        setViewer(loaded)
      } catch (reason: unknown) {
        if (!cancelled) setError(`Could not open the PDF: ${errorMessage(reason)}`)
      }
    })()

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
    const generation = ++renderGenerationRef.current
    const selectedPage = clampPdfPage(page, pageCount)
    const render = renderQueueRef.current
      .catch(() => undefined)
      .then(async () => {
        const canvas = canvasRef.current
        if (!canvas || cancelled || viewerRef.current !== viewer) return
        await renderPageToCanvas(viewer, selectedPage, canvas, 1)
        if (cancelled || generation !== renderGenerationRef.current || viewerRef.current !== viewer) return
      })
      .catch((reason: unknown) => {
        if (!cancelled && generation === renderGenerationRef.current) {
          setError(`Could not render page ${selectedPage}: ${errorMessage(reason)}`)
        }
      })
    renderQueueRef.current = render
    return () => {
      cancelled = true
    }
  }, [page, pageCount, viewer])

  useEffect(() => {
    if (!manager) return
    manager.publish({ page, annotLocalId: selectedAnnot ? `o${selectedAnnot.objNum}` : null })
  }, [manager, page, selectedAnnot])

  const disconnect = () => {
    managerRef.current?.stop()
    transportRef.current?.close()
    managerRef.current = null
    transportRef.current = null
    engineRef.current = null
    setManager(null)
    setPeers([])
    setMetrics({ peers: 0, applied: 0, pending: pendingRef.current.length })
  }

  const submitLocal = async (operation: PdfCollabOperation) => {
    if (!engineRef.current) return
    const encoded = encodePdfCollabOperation(operation)
    setError('')
    try {
      await runExclusive(async () => {
        const current = bytesRef.current
        if (!current) throw new Error('PDF is still loading.')
        const next = await applyPdfCollabOps(current, [encoded])
        bytesRef.current = next
        setBytes(next)
        pendingRef.current = [...pendingRef.current, encoded]
        engineRef.current?.submitSoon()
        refreshMetrics()
      })
      setStatus(
        connection === 'live'
          ? 'Annotation queued for the ordered log.'
          : 'Annotation applied locally. Join a room to sync it.',
      )
    } catch (cause) {
      setError(errorMessage(cause))
    }
  }

  const joinRoom = async (path: string) => {
    setBusy(true)
    setConnection('joining')
    setError('')
    disconnect()
    pendingRef.current = []
    try {
      const seed = seedRef.current ?? await makeCollabPdfSample()
      seedRef.current = seed
      bytesRef.current = seed
      setBytes(seed)
      const transport = await createHttpCollabTransport(COLLAB_API_BASE)
      transportRef.current = transport
      const nextManager = new PdfPresenceManager(transport as unknown as CollabTransport<PdfSelection>, { path, name })
      managerRef.current = nextManager
      nextManager.onChange(() => refreshMetrics(nextManager))
      nextManager.onFileChanged(() => {
        setStatus('The backing artifact changed outside this editor. Remint the room before continuing.')
      })
      await nextManager.start()
      if (!nextManager.clientId) throw new Error('The sidecar did not join the collaboration room.')
      const engine = await nextManager.startSync({
        applyRemote(ops) {
          return runExclusive(async () => {
            const current = bytesRef.current
            if (!current) throw new Error('PDF is still loading.')
            const next = await applyPdfCollabOps(current, ops)
            bytesRef.current = next
            setBytes(next)
            refreshMetrics(nextManager)
          })
        },
        getPending() {
          const ops = pendingRef.current
          return ops.length > 0 ? { ops: ops.slice() } : null
        },
        onAcked(ops) {
          pendingRef.current = pendingRef.current.slice(ops.length)
          refreshMetrics(nextManager)
        },
        onResync() {
          setConnection('error')
          setError('The operation log requires a PDF reload. Remint the artifact to restart this proof.')
        },
      })
      if (!engine) throw new Error('The sidecar did not start PDF annotation sync.')
      engineRef.current = engine
      await runExclusive(async () => undefined)
      setManager(nextManager)
      setArtifact(path)
      refreshMetrics(nextManager)
      setConnection('live')
      setStatus(`Live as ${nextManager.state.me?.name ?? name}. Add a highlight or form value, then watch it appear in the second tab.`)
    } catch (cause) {
      disconnect()
      setConnection('error')
      setError(errorMessage(cause))
    } finally {
      setBusy(false)
    }
  }

  const onMint = async () => {
    setBusy(true)
    setError('')
    try {
      const pdf = bytesRef.current ?? seedRef.current ?? await makeCollabPdfSample()
      seedRef.current ??= pdf
      const id = await mintCollabPdfArtifact(pdf)
      setArtifact(id)
      await joinRoom(id)
    } catch (cause) {
      setConnection('error')
      setError(errorMessage(cause))
      setBusy(false)
    }
  }

  const note = annots.find((annot) => annot.subtype === 'note')
  const ready = viewer !== null && pageCount > 0
  const syncing = connection === 'live' && !busy && engineRef.current !== null

  return (
    <>
      <div className="native-toolbar collab-toolbar workbench-toolbar" role="group" aria-label="PDF collaboration actions">
        <strong>Live PDF annotation protocol proof</strong>
        <span className={`collab-connection collab-connection--${connection} workbench-badge`} role="status" aria-label={`Connection: ${connection}`}>
          <i aria-hidden="true" />
          {connectionLabel(connection)}
        </span>
        <label>
          Name
          <input value={name} onChange={(event) => setName(event.target.value)} disabled={connection === 'live'} />
        </label>
        <label>
          Artifact
          <input
            value={artifact}
            onChange={(event) => setArtifact(event.target.value.trim())}
            placeholder="art_…"
            spellCheck={false}
            disabled={busy}
          />
        </label>
        <button type="button" className="workbench-button workbench-button--primary" disabled={busy || !artifact} onClick={() => void joinRoom(artifact)}>
          Join
        </button>
        <button type="button" className="workbench-button" disabled={busy} onClick={() => void onMint()}>
          Mint sample
        </button>
        <button type="button" className="workbench-button" disabled={!artifact} onClick={() => window.open(window.location.href, '_blank')}>
          Open second tab
        </button>
        {manager ? <PresenceStack manager={manager} /> : null}
      </div>

      <p className="native-status workbench-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : connection}>{status}</p>
      {error ? <p className="native-error workbench-callout workbench-callout--error" role="alert">{error}</p> : null}

      <section className="collab-boundaries workbench-boundary" aria-labelledby="collab-pdf-boundaries-title">
        <div>
          <h2 id="collab-pdf-boundaries-title">What this proves</h2>
          <p>{PDF_PROOF}</p>
        </div>
        <div className="collab-security workbench-callout workbench-callout--warning" role="note">
          <strong>Local development only</strong>
          <p>{COLLAB_COPY.security}</p>
        </div>
        <p className="collab-completion">{COLLAB_COPY.completion}</p>
      </section>

      <div className="collab-body collab-body--univer">
        <div className="collab-univer-wrap">
          <div className="collab-univer-heading">
            <div>
              <h2>Shared PDF</h2>
              <p>Annotation and form values sync. Page, text, and image edits are not collaborated.</p>
            </div>
            <span>{ready ? `Page ${page} of ${pageCount}` : 'Loading PDF'}</span>
          </div>
          <div className="split" style={{ flex: 1, minHeight: 500 }}>
            <div className="split-main pages pdf-pages">
              <div className="toolstrip pdf-toolbar workbench-toolbar" role="group" aria-label="PDF annotation actions">
                <button type="button" className="workbench-button" disabled={!ready || page <= 1} onClick={() => setPage((current) => clampPdfPage(current - 1, pageCount))}>Previous</button>
                <label className="pdf-page-input">
                  Page
                  <span className="pdf-page-input__control">
                    <input type="number" min={1} max={Math.max(1, pageCount)} value={page} disabled={!ready} onChange={(event) => setPage(clampPdfPage(Number(event.target.value), pageCount))} />
                    <span>of {pageCount || '—'}</span>
                  </span>
                </label>
                <button type="button" className="workbench-button" disabled={!ready || page >= pageCount} onClick={() => setPage((current) => clampPdfPage(current + 1, pageCount))}>Next</button>
                <button
                  type="button"
                  className="workbench-button workbench-button--primary"
                  disabled={!syncing}
                  onClick={() => void submitLocal({
                    kind: 'annotation.markup',
                    value: { page, type: 'highlight', color: [1, 0.92, 0.2], quads: SAMPLE_HIGHLIGHT_QUADS },
                  })}
                >
                  Highlight
                </button>
              </div>
              {!ready && !error ? <p className="pdf-empty">Loading PDF…</p> : null}
              <canvas ref={canvasRef} className="pdf-canvas" aria-label={ready ? `Rendered page ${page} of ${pageCount}` : 'PDF page'} />
            </div>
            <aside className="split-side pdf-inspector workbench-inspector" aria-label="PDF annotation controls">
              <div className="ioc-panel">
                <strong>Sticky note</strong>
                <label>
                  Contents
                  <textarea
                    value={noteDraft}
                    onChange={(event) => {
                      noteDirtyRef.current = true
                      setNoteDraft(event.target.value)
                    }}
                    rows={3}
                    disabled={!note || !syncing}
                  />
                </label>
                <button
                  type="button"
                  className="workbench-button"
                  disabled={!note || !syncing}
                  onClick={() => {
                    if (!note) return
                    noteDirtyRef.current = false
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
                >
                  Save note
                </button>
              </div>
              <div className="ioc-panel">
                <strong>Form values</strong>
                <label>
                  Memo
                  <input
                    value={memo}
                    onChange={(event) => {
                      memoDirtyRef.current = true
                      setMemo(event.target.value)
                    }}
                    disabled={!syncing}
                  />
                </label>
                <label>
                  <input
                    type="checkbox"
                    checked={agree}
                    disabled={!syncing}
                    onChange={(event) => {
                      agreeDirtyRef.current = true
                      setAgree(event.target.checked)
                    }}
                  />
                  Agree
                </label>
                <button
                  type="button"
                  className="workbench-button"
                  disabled={!syncing}
                  onClick={() => {
                    memoDirtyRef.current = false
                    agreeDirtyRef.current = false
                    void (async () => {
                      await submitLocal({ kind: 'form.value', value: { name: MEMO_FIELD, kind: 'text', value: memo } })
                      await submitLocal({ kind: 'form.value', value: { name: AGREE_FIELD, kind: 'checkbox', checked: agree } })
                    })()
                  }}
                >
                  Apply form values
                </button>
              </div>
              <div className="ioc-panel">
                <strong>Annotations</strong>
                {annots.length === 0 ? <p>No markup or notes yet.</p> : (
                  <ol className="pdf-result-list">
                    {annots.map((annot) => (
                      <li key={`${annot.page}-${annot.objNum}-${annot.subtype}`}>
                        <button
                          type="button"
                          className="workbench-button workbench-button--quiet"
                          onClick={() => {
                            setSelectedAnnot(annot)
                            setPage(clampPdfPage(annot.page, pageCount))
                          }}
                        >
                          Page {annot.page} · {annot.subtype} · {annot.objNum}
                        </button>
                        <button
                          type="button"
                          className="workbench-button"
                          disabled={!syncing}
                          onClick={() => {
                            setSelectedAnnot(annot)
                            void submitLocal({
                              kind: 'annotation.delete',
                              value: {
                                page: annot.page,
                                objNum: annot.objNum,
                                subtype: annot.subtype,
                                rect: annot.rect,
                                contents: annot.subtype === 'note' ? annot.contents : undefined,
                              },
                            })
                          }}
                        >
                          Delete
                        </button>
                      </li>
                    ))}
                  </ol>
                )}
              </div>
            </aside>
          </div>
        </div>
        <aside className="native-side collab-protocol-side workbench-inspector">
          <h3>Protocol state</h3>
          <dl className="collab-metrics">
            <div><dt>Peers</dt><dd>{metrics.peers}</dd></div>
            <div><dt>Applied sequence</dt><dd>{metrics.applied}</dd></div>
            <div><dt>Pending operations</dt><dd>{metrics.pending}</dd></div>
          </dl>
          <h4>Room artifact</h4>
          <p className="native-muted">{artifact || 'No artifact yet.'}</p>
          <h4>Peer pages</h4>
          {peers.length === 0 ? <p className="native-muted">No remote peers yet.</p> : (
            <ul>
              {peers.map((peer) => (
                <li key={peer.client_id}>
                  <span className="collab-dot" style={{ background: peer.color }} />
                  {peer.name}
                  {peer.selection ? ` · page ${peer.selection.page}` : ''}
                  {peer.selection?.annotLocalId ? ` · ${peer.selection.annotLocalId}` : ''}
                </li>
              ))}
            </ul>
          )}
          <h4>Try it</h4>
          <ol>
            <li>Mint the sample PDF. It is only an opaque room key.</li>
            <li>Open the same URL in another tab and join.</li>
            <li>Add a highlight, edit the sticky note, or fill the form.</li>
            <li>Confirm the ordered annotation and remote page presence in both tabs.</li>
          </ol>
        </aside>
      </div>
    </>
  )
}

export const CollabPdfPlaceholder = CollabPdfPanel
export default CollabPdfPanel
