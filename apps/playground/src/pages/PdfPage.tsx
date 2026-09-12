import { useEffect, useRef, useState, type ChangeEvent, type FormEvent } from 'react'
import pdfWorkerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { getPdfLoadOptions } from 'virtual:injoffice-pdf-resources'
import {
  DsButton,
  DsCallout,
  DsField,
  DsInput,
  DsSelect,
} from '../design-system/primitives'
import '../design-system/live-create-edit.css'
import type { MarkupType } from '../../../../packages/pdf/src/annotate/types'
import type { TextAppearanceFont } from '../../../../packages/pdf/src/annotate/forms'
import type { PdfDocumentInfo } from '../../../../packages/pdf/src/types'
import {
  configurePdfWorker,
  PdfViewerDocument,
  renderPageToCanvas,
  type SearchMatch,
  type PdfOutlineItem,
} from '../../../../packages/pdf/src/viewer'
import { PDF_NODE_PREFIX } from '../pdfHostRoutes'
import { createPdfDemoFixture, PDF_DEMO_FILE_NAME } from '../pdfDemoFixture'
import {
  applyPdfAnnotDelete,
  applyPdfPlacedDrawing,
  applyPdfFormValues,
  pdfFormResultMessage,
  restorePdfSkippedDrafts,
  applyPdfMarkup,
  applyPdfNote,
  applyPdfNoteEdit,
  applyPdfPageOp,
  applyPdfStamp,
  listPdfAnnots,
  listPdfFormFields,
  mergePdfBytes,
  movePageOrder,
  pdfGeometry,
  splitPdfPrefix,
  type PdfAnnot,
  type PdfFormField,
} from '../pdfWorkbench'
import { recordPdfEdit, travelPdfHistory, type PdfHistory } from '../pdfInteraction'
import { PdfInteractionLayer, type PdfPlacementTool } from './PdfInteractionLayer'
import { PdfTextLayer } from './PdfTextLayer'
import { PdfContinuousView } from './PdfContinuousView'

configurePdfWorker(pdfWorkerUrl)

type OutlineItem = PdfOutlineItem
type Inspector = 'inspect' | 'pages' | 'mark' | 'draw' | 'forms' | 'stamp' | 'host'

const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2]
const INSPECTORS: { id: Inspector; label: string }[] = [
  { id: 'inspect', label: 'Inspect' },
  { id: 'pages', label: 'Pages' },
  { id: 'mark', label: 'Mark' },
  { id: 'draw', label: 'Draw' },
  { id: 'forms', label: 'Forms' },
  { id: 'stamp', label: 'Stamp' },
  { id: 'host', label: 'Advanced' },
]

function OutlineList({ items, onSelect }: { items: OutlineItem[]; onSelect: (page: number) => void }) {
  return <ol className="pdf-result-list">{items.map((item, index) => <li key={`${item.title}-${index}`}>
    <button type="button" className="workbench-button workbench-button--quiet" disabled={item.pageIndex === null} onClick={() => item.pageIndex !== null && onSelect(item.pageIndex)}>{item.title}{item.pageIndex === null ? '' : ` · ${item.pageIndex}`}</button>
    {item.children.length > 0 && <OutlineList items={item.children} onSelect={onSelect} />}
  </li>)}</ol>
}

export function clampPdfPage(page: number, pageCount: number): number {
  if (pageCount < 1 || !Number.isFinite(page)) return 1
  return Math.max(1, Math.min(Math.trunc(page), pageCount))
}

export function pdfDownloadName(fileName: string, edited: boolean): string {
  const base = fileName.replace(/\.pdf$/i, '').replace(/[\\/:*?"<>|]/g, '-').trim().replace(/^\.+/, '') || 'document'
  return `${base}${edited ? '-edited' : ''}.pdf`
}

function errorMessage(error: unknown): string {
  return error instanceof Error && error.message ? error.message : 'Unknown PDF error'
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  bytes.forEach((value) => {
    binary += String.fromCharCode(value)
  })
  return btoa(binary)
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

export default function PdfPage() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const uploadRef = useRef<HTMLInputElement | null>(null)
  const mergeRef = useRef<HTMLInputElement | null>(null)
  const viewerRef = useRef<PdfViewerDocument | null>(null)
  const renderQueueRef = useRef<Promise<void>>(Promise.resolve())
  const renderGenerationRef = useRef(0)
  const searchControllerRef = useRef<AbortController | null>(null)
  const [bytes, setBytes] = useState<Uint8Array | null>(null)
  const [fileName, setFileName] = useState(PDF_DEMO_FILE_NAME)
  const [edited, setEdited] = useState(false)
  const [viewer, setViewer] = useState<PdfViewerDocument | null>(null)
  const [page, setPage] = useState(1)
  const [pageCount, setPageCount] = useState(0)
  const [zoom, setZoom] = useState(1)
  const [continuous, setContinuous] = useState(false)
  const [info, setInfo] = useState('Preparing the bundled sample…')
  const [text, setText] = useState('')
  const [outline, setOutline] = useState<OutlineItem[]>([])
  const [query, setQuery] = useState('')
  const [matches, setMatches] = useState<SearchMatch[]>([])
  const [searched, setSearched] = useState(false)
  const [searching, setSearching] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState<'loading' | 'rendering' | 'editing' | null>('loading')
  const [geometry, setGeometry] = useState<PdfDocumentInfo | null>(null)
  const [annots, setAnnots] = useState<PdfAnnot[]>([])
  const [fields, setFields] = useState<PdfFormField[]>([])
  const [fieldBytes, setFieldBytes] = useState<Uint8Array | null>(null)
  const [formNotice, setFormNotice] = useState<string | null>(null)
  const [formAppearanceFont, setFormAppearanceFont] = useState<TextAppearanceFont | 'viewer'>('viewer')
  const [unsavedFormNames, setUnsavedFormNames] = useState<string[]>([])
  const pendingFormDrafts = useRef<{ bytes: Uint8Array; drafts: PdfFormField[]; skipped: { name: string }[] } | null>(null)
  const [noteText, setNoteText] = useState('Shared note')
  const [hostOld, setHostOld] = useState('InjOffice')
  const [hostNew, setHostNew] = useState('InjOffice PDF')
  const [hostNote, setHostNote] = useState('Node host idle. Available during npm run dev.')
  const [inspector, setInspector] = useState<Inspector>('pages')
  const [history, setHistory] = useState<PdfHistory>({ past: [], future: [] })
  const [selection, setSelection] = useState<number[][]>([])
  const [tool, setTool] = useState<PdfPlacementTool>(null)
  const [searchTerm, setSearchTerm] = useState('')

  useEffect(() => {
    searchControllerRef.current?.abort()
    setSearching(false)
    setSearched(false)
    setMatches([])
    setSearchTerm('')
    return () => { searchControllerRef.current?.abort() }
  }, [query, bytes])

  useEffect(() => { setSelection([]); setTool(null) }, [page, zoom, bytes, inspector])
  useEffect(() => {
    if (!tool) return
    const cancel = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && canvasRef.current?.getClientRects().length && !canvasRef.current.closest('[inert]')) setTool(null)
    }
    window.addEventListener('keydown', cancel)
    return () => window.removeEventListener('keydown', cancel)
  }, [tool])

  const applyBytes = async (next: Uint8Array, message: string) => {
    if (bytes) setHistory(current => recordPdfEdit(current, { bytes, page, edited }))
    setBytes(next)
    setEdited(true)
    setInfo(message)
  }

  const travel = (direction: 'undo' | 'redo') => {
    if (!bytes || busy) return
    const result = travelPdfHistory(history, { bytes, page, edited }, direction)
    if (!result) return
    setFormNotice(null)
    setHistory(result.history)
    setPage(result.snapshot.page)
    setEdited(result.snapshot.edited)
    setBytes(result.snapshot.bytes)
  }

  const runEdit = async (work: () => Promise<Uint8Array>, message: string) => {
    if (!bytes || busy) return
    setFormNotice(null)
    setBusy('editing')
    setError(null)
    try {
      const next = await work()
      await applyBytes(next, message)
    } catch (reason: unknown) {
      setError(errorMessage(reason))
    } finally {
      setBusy(null)
    }
  }

  const applyFormDrafts = async () => {
    if (!bytes || busy || fieldBytes !== bytes) return
    setBusy('editing')
    setError(null)
    setInfo('')
    setFormNotice(null)
    try {
      const result = await applyPdfFormValues(bytes, fields, formAppearanceFont === 'viewer' ? undefined : { textAppearance: { font: formAppearanceFont } })
      const message = pdfFormResultMessage(result)
      setFormNotice(message)
      setUnsavedFormNames(result.skipped.map(({ name }) => name))
      if (result.applied > 0) {
        pendingFormDrafts.current = { bytes: result.bytes, drafts: fields, skipped: result.skipped }
        await applyBytes(result.bytes, message)
      }
      else setInfo(message)
    } catch (reason: unknown) {
      setError(errorMessage(reason))
    } finally {
      setBusy(null)
    }
  }

  const callNodeHost = async (route: string, extra: Record<string, unknown> = {}) => {
    if (!bytes) throw new Error('Open a PDF first.')
    let response: Response
    try {
      response = await fetch(`${PDF_NODE_PREFIX}/${route}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ bytes: bytesToBase64(bytes), page, ...extra }),
      })
    } catch {
      throw new Error('Node PDF host is unreachable. Run the playground with npm run dev.')
    }
    const payload = await response.json() as { error?: string; bytes?: string; png?: string; images?: unknown; skipped?: unknown; applied?: unknown; proofs?: unknown; pagesOcred?: unknown; failed?: unknown }
    if (response.status === 404) throw new Error('Node PDF host is only available during npm run dev, not a static Pages build.')
    if (!response.ok) throw new Error(payload.error || `PDF host failed (${response.status})`)
    if (payload.bytes) {
      await applyBytes(base64ToBytes(payload.bytes), `Node host ${route} applied.`)
      setHostNote(JSON.stringify({ applied: payload.applied, skipped: payload.skipped, proofs: payload.proofs, pagesOcred: payload.pagesOcred, failed: payload.failed }, null, 2))
      return
    }
    if (payload.png) {
      const link = document.createElement('a')
      link.href = `data:image/png;base64,${payload.png}`
      link.download = `page-${page}.png`
      link.click()
      setHostNote(`Rendered page ${page} PNG (${payload.png.length} b64 chars).`)
      return
    }
    setHostNote(JSON.stringify(payload, null, 2))
  }

  const runHost = (route: string, extra: Record<string, unknown> = {}) => {
    if (!bytes || busy) return
    setBusy('editing')
    setError(null)
    void callNodeHost(route, extra)
      .catch((reason: unknown) => setError(errorMessage(reason)))
      .finally(() => setBusy(null))
  }

  useEffect(() => {
    let cancelled = false
    void createPdfDemoFixture()
      .then((fallback) => { if (!cancelled) setBytes(fallback) })
      .catch((reason: unknown) => {
        if (cancelled) return
        setError(`Could not generate the bundled PDF: ${errorMessage(reason)}`)
        setInfo('PDF sample unavailable')
        setBusy(null)
      })
    return () => { cancelled = true }
  }, [])

  useEffect(() => {
    if (!bytes) return
    let cancelled = false
    let ownedViewer: PdfViewerDocument | null = null
    renderGenerationRef.current += 1
    viewerRef.current = null
    setViewer(null)
    setPageCount(0)
    setText('')
    setOutline([])
    setMatches([])
    setSearched(false)
    setSearching(false)
    setError(null)
    setBusy('loading')
    setInfo(`Opening ${fileName}…`)
    const canvas = canvasRef.current
    if (canvas) {
      canvas.width = 0
      canvas.height = 0
      canvas.style.width = '0px'
      canvas.style.height = '0px'
    }

    void (async () => {
      try {
        const loaded = await PdfViewerDocument.load(bytes, getPdfLoadOptions())
        ownedViewer = loaded
        if (cancelled) { await loaded.destroy(); return }
        viewerRef.current = loaded
        setPageCount(loaded.pageCount)
        setPage((current) => clampPdfPage(current, loaded.pageCount))
        setViewer(loaded)
        setBusy(null)
        setInfo(`${fileName} · ${loaded.pageCount} page${loaded.pageCount === 1 ? '' : 's'}`)
        const [nextOutline, nextGeometry, nextAnnots, nextFields] = await Promise.all([
          loaded.getOutlineTree().catch(() => [] as OutlineItem[]),
          pdfGeometry(bytes),
          listPdfAnnots(bytes),
          listPdfFormFields(bytes),
        ])
        if (!cancelled && viewerRef.current === loaded) {
          setOutline(nextOutline)
          setGeometry(nextGeometry)
          setAnnots(nextAnnots)
          const pending = pendingFormDrafts.current
          if (pending?.bytes === bytes) {
            setFields(restorePdfSkippedDrafts(nextFields, pending.drafts, pending.skipped))
            setUnsavedFormNames(pending.skipped.map(({ name }) => name))
          } else {
            setFields(nextFields)
            setUnsavedFormNames([])
          }
          pendingFormDrafts.current = null
          setFieldBytes(bytes)
        }
      } catch (reason: unknown) {
        if (cancelled) return
        setError(`Could not open the PDF: ${errorMessage(reason)}`)
        setInfo('PDF viewer unavailable')
        setBusy(null)
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
  }, [bytes, fileName])

  useEffect(() => {
    if (!viewer || pageCount < 1) return
    if (continuous) {
      setBusy(null)
      setInfo(`${fileName} · page ${page} of ${pageCount} · continuous reading · ${Math.round(zoom * 100)}%`)
      let cancelled = false
      void viewer.getPageText(page).then(value => { if (!cancelled) setText(value) }).catch(() => { if (!cancelled) setText('') })
      return () => { cancelled = true }
    }
    let cancelled = false
    const controller = new AbortController()
    const generation = ++renderGenerationRef.current
    const selectedPage = clampPdfPage(page, pageCount)
    setBusy('rendering')
    const render = renderQueueRef.current
      .catch(() => undefined)
      .then(async () => {
        if (cancelled || viewerRef.current !== viewer) return
        const size = await viewer.getPageSize(selectedPage)
        const extractedText = await viewer.getPageText(selectedPage)
        const canvas = canvasRef.current
        if (!canvas || cancelled || viewerRef.current !== viewer) return
        await renderPageToCanvas(viewer, selectedPage, canvas, zoom, undefined, { signal: controller.signal })
        if (cancelled || generation !== renderGenerationRef.current || viewerRef.current !== viewer) return
        setText(extractedText)
        setInfo(`${fileName} · page ${selectedPage} of ${pageCount} · ${Math.round(size.width)}×${Math.round(size.height)} pt · ${Math.round(zoom * 100)}%`)
        setBusy(null)
      })
      .catch((reason: unknown) => {
        if (cancelled || generation !== renderGenerationRef.current) return
        setError(`Could not render page ${selectedPage}: ${errorMessage(reason)}`)
        setBusy(null)
      })
    renderQueueRef.current = render
    return () => { cancelled = true; controller.abort() }
  }, [continuous, fileName, page, pageCount, viewer, zoom])

  const openPdf = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0]
    event.currentTarget.value = ''
    if (!file) return
    if (edited && !window.confirm('Open another PDF and discard your current edits? Download your edited PDF first if you want to keep it.')) return
    setError(null)
    setFormNotice(null)
    setBusy('loading')
    try {
      const nextBytes = new Uint8Array(await file.arrayBuffer())
      if (nextBytes.length === 0) throw new Error('the selected file is empty')
      // Keep the current document and undo stack until the replacement is usable.
      const candidate = await PdfViewerDocument.load(nextBytes, getPdfLoadOptions())
      try { await candidate.getPage(1) } finally { await candidate.destroy() }
      setFileName(file.name || 'document.pdf')
      setFormAppearanceFont('viewer')
      setEdited(false)
      setHistory({ past: [], future: [] })
      setPage(1)
      setBytes(nextBytes)
    } catch (reason: unknown) {
      setError(`Could not read the selected PDF: ${errorMessage(reason)}`)
    } finally {
      setBusy(null)
    }
  }

  const mergePdf = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0]
    event.currentTarget.value = ''
    if (!file || !bytes) return
    await runEdit(async () => mergePdfBytes(bytes, new Uint8Array(await file.arrayBuffer())), `Merged ${file.name}.`)
  }

  const downloadPdf = (payload = bytes, name = fileName) => {
    if (!payload) return
    const copy = Uint8Array.from(payload)
    const url = URL.createObjectURL(new Blob([copy.buffer], { type: 'application/pdf' }))
    const link = document.createElement('a')
    link.href = url
    link.download = pdfDownloadName(name, edited)
    link.click()
    window.setTimeout(() => URL.revokeObjectURL(url), 0)
  }

  const search = async (event: FormEvent) => {
    event.preventDefault()
    const activeViewer = viewerRef.current
    if (!activeViewer || query.trim().length === 0) return
    searchControllerRef.current?.abort()
    const controller = new AbortController()
    searchControllerRef.current = controller
    setSearching(true)
    setError(null)
    try {
      const nextMatches = await activeViewer.search(query.trim(), { signal: controller.signal })
      if (controller.signal.aborted || viewerRef.current !== activeViewer) return
      setMatches(nextMatches)
      setSearchTerm(query.trim())
      if (nextMatches[0]) setPage(nextMatches[0].pageIndex)
      setSearched(true)
      setSearching(false)
    } catch (reason: unknown) {
      if (controller.signal.aborted || viewerRef.current !== activeViewer) return
      setError(`Could not search the PDF: ${errorMessage(reason)}`)
      setSearching(false)
    }
  }

  const ready = viewer !== null && pageCount > 0
  const locked = !ready || busy !== null
  const current = geometry?.pages[page - 1]

  return (
    <div className="platen-fill workbench-surface ds" data-demo-surface="pdf" data-demo-dirty={edited} data-demo-busy={busy === 'editing'}>
      <div className="toolstrip pdf-toolbar workbench-toolbar ds-workstrip" role="group" aria-label="PDF actions">
        <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={busy !== null} onClick={() => uploadRef.current?.click()}>Open PDF</DsButton>
        <input ref={uploadRef} className="visually-hidden" type="file" accept="application/pdf,.pdf" aria-label="Open a PDF file" onChange={(event) => { void openPdf(event) }} />
        <DsButton className="workbench-button" disabled={!ready || busy !== null || page <= 1} onClick={() => setPage((currentPage) => clampPdfPage(currentPage - 1, pageCount))}>Previous</DsButton>
        <span className="pdf-page-input">
          Page
          <span className="pdf-page-input__control">
            <input aria-label="Page number" type="number" min={1} max={Math.max(1, pageCount)} value={page} disabled={locked} onChange={(event) => setPage(clampPdfPage(Number(event.target.value), pageCount))} />
            <span>of {pageCount || '—'}</span>
          </span>
        </span>
        <DsButton className="workbench-button" disabled={!ready || busy !== null || page >= pageCount} onClick={() => setPage((currentPage) => clampPdfPage(currentPage + 1, pageCount))}>Next</DsButton>
        <DsField label="Zoom">
          <DsSelect aria-label="Zoom" value={zoom} disabled={!ready} onChange={(event) => setZoom(Number(event.target.value))}>{ZOOM_STEPS.map((step) => <option key={step} value={step}>{Math.round(step * 100)}%</option>)}</DsSelect>
        </DsField>
        <DsButton disabled={!ready || busy === 'editing'} aria-pressed={continuous} onClick={() => { setTool(null); setSelection([]); setContinuous(value => !value); if (!continuous) setInspector('inspect') }}>{continuous ? 'Single page' : 'Continuous reading'}</DsButton>
        <DsButton variant="outlined" className="workbench-button" disabled={!bytes || busy === 'loading'} onClick={() => downloadPdf()}>Download {edited ? 'edited PDF' : 'PDF'}</DsButton>
        <DsButton disabled={locked || !history.past.length} onClick={() => travel('undo')}>Undo</DsButton>
        <DsButton disabled={locked || !history.future.length} onClick={() => travel('redo')}>Redo</DsButton>
      </div>
      <p className="native-status workbench-status ds-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : busy ?? 'ready'}>
        {busy === 'rendering' ? 'Rendering…' : busy === 'editing' ? 'Applying PDF operation…' : info}
      </p>
      {error && <DsCallout tone="refuse" title="PDF error">{error}</DsCallout>}
      {tool && <p className="pdf-placement-hint" role="status">{tool === 'note' ? 'Click the page to place your note.' : `Drag on the page to draw ${tool === 'ink' ? 'an ink stroke' : `a ${tool}`}.`} Keyboard: Tab to the page, arrow keys to position, Enter to place each endpoint. <DsButton onClick={() => setTool(null)}>Cancel drawing</DsButton></p>}
      <div className="split ds-split">
        {continuous && viewer ? <div className="split-main ds-split-main pdf-reader-host"><PdfContinuousView key={fileName + ':' + pageCount} viewer={viewer} page={page} zoom={zoom} onPageChange={setPage} /></div> : <div className="split-main pages pdf-pages ds-split-main" aria-busy={busy === 'loading' || busy === 'rendering' || busy === 'editing'}>
          <article className="ds-pdf-sheet">
            {!ready && !error && <p className="pdf-empty ds-muted">Loading PDF…</p>}
            <div className="pdf-page-stage">
              <canvas ref={canvasRef} className="pdf-canvas" aria-label={ready ? `Rendered page ${page} of ${pageCount}` : 'PDF page'} />
              {viewer && <PdfTextLayer viewer={viewer} page={page} zoom={zoom} disabled={locked || inspector === 'mark' || tool !== null} />}
              {viewer && <PdfInteractionLayer key={`${page}-${zoom}`} viewer={viewer} page={page} zoom={zoom} locked={locked} selectText={inspector === 'mark'} tool={tool} matches={matches} queryLength={searchTerm.length} onSelection={setSelection} onPlace={(kind, points) => {
                setTool(null)
                void runEdit(() => kind === 'note' ? applyPdfNote(bytes!, page, noteText, points[0]) : applyPdfPlacedDrawing(bytes!, page, kind, points), `Added ${kind} on page ${page}.`)
              }} />}
            </div>
          </article>
        </div>}
        <aside className="split-side pdf-inspector workbench-inspector ds-split-side" aria-label="PDF operations">
          <div className="view-switcher" role="tablist" aria-label="PDF tool panels">
            {INSPECTORS.map((item) => (
              <button key={item.id} id={`pdf-tab-${item.id}`} type="button" role="tab" className="ds-pick" aria-selected={inspector === item.id} aria-controls="pdf-operation-panel" tabIndex={inspector === item.id ? 0 : -1} onClick={() => { setInspector(item.id); if (item.id !== 'inspect') setContinuous(false) }} onKeyDown={event => {
                const index = INSPECTORS.findIndex(entry => entry.id === item.id)
                const next = event.key === 'ArrowRight' ? (index + 1) % INSPECTORS.length : event.key === 'ArrowLeft' ? (index + INSPECTORS.length - 1) % INSPECTORS.length : event.key === 'Home' ? 0 : event.key === 'End' ? INSPECTORS.length - 1 : -1
                if (next < 0) return
                event.preventDefault()
                setInspector(INSPECTORS[next]!.id)
                if (INSPECTORS[next]!.id !== 'inspect') setContinuous(false)
                document.getElementById(`pdf-tab-${INSPECTORS[next]!.id}`)?.focus()
              }}>{item.label}</button>
            ))}
          </div>

          <div id="pdf-operation-panel" className="pdf-operation-panel" role="tabpanel" aria-labelledby={`pdf-tab-${inspector}`} tabIndex={0}>
          {inspector === 'inspect' && (
            <div className="ioc-panel ds-panel pdf-search">
              <span className="ds-eyebrow">Search document</span>
              <form onSubmit={(event) => { void search(event) }}>
                <DsField label="Query">
                  <DsInput aria-label="Search PDF text" type="search" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Find extracted text" />
                </DsField>
                <div className="pdf-search-row">
                  <DsButton type="submit" className="workbench-button" disabled={!ready || searching || query.trim().length === 0}>{searching ? 'Finding…' : 'Find'}</DsButton>
                </div>
              </form>
              {searched && <p className="ds-muted" role="status">{matches.length} match{matches.length === 1 ? '' : 'es'}. {continuous ? 'Select a result to jump to its page. Switch to Single page to see passage highlights.' : 'Matching text passages are highlighted on the page.'}</p>}
              {matches.length > 0 && <ol className="pdf-result-list">{matches.map((match, index) => <li key={`${match.pageIndex}-${match.offset}-${index}`}><button type="button" className="workbench-button workbench-button--quiet" onClick={() => setPage(match.pageIndex)}><span>Page {match.pageIndex}</span>{match.snippet || query}</button></li>)}</ol>}
              <p className="ds-muted">Geometry · {geometry ? `${geometry.pageCount} pages` : '—'}{current ? ` · ${Math.round(current.width)}×${Math.round(current.height)} pt · ${current.rotation}°` : ''}</p>
              <span className="ds-eyebrow">Document outline</span>
              {outline.length === 0 ? <p className="ds-muted">No outline entries.</p> : <OutlineList items={outline} onSelect={setPage} />}
              <span className="ds-eyebrow">Page {page} extracted text</span>
              <p className="ds-muted">{text || 'No extractable text on this page.'}</p>
            </div>
          )}

          {inspector === 'pages' && (
            <div className="ioc-panel ds-panel">
              <span className="ds-eyebrow">Page operations</span>
              <p className="ds-muted">Rearrange pages or combine documents. Undo restores your last 20 changes.</p>
              <div className="ds-stack">
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'rotate', pages: [page], degrees: 90 }), `Rotated page ${page} by 90°.`)}>Rotate 90°</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'rotate', pages: [page], degrees: 180 }), `Rotated page ${page} by 180°.`)}>Rotate 180°</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked || pageCount < 2} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'delete', pages: [page] }), `Deleted page ${page}.`)}>Delete page</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'insertBlank', at: page + 1 }), `Inserted a blank page after ${page}.`)}>Insert blank after</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked || !movePageOrder(pageCount, page, -1)} onClick={() => { const order = movePageOrder(pageCount, page, -1); if (order) void runEdit(() => applyPdfPageOp(bytes!, { type: 'reorder', order }), 'Moved page earlier.') }}>Move earlier</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked || !movePageOrder(pageCount, page, 1)} onClick={() => { const order = movePageOrder(pageCount, page, 1); if (order) void runEdit(() => applyPdfPageOp(bytes!, { type: 'reorder', order }), 'Moved page later.') }}>Move later</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked || !current} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'crop', pages: [page], box: { x: 36, y: 36, width: Math.max(72, current!.width - 72), height: Math.max(72, current!.height - 72) } }), `Cropped page ${page}.`)}>Crop 36pt inset</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'resize', pages: [page], width: 612, height: 792, fit: 'contain' }), `Resized page ${page} to letter.`)}>Resize to letter</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'nUp', n: 2 }), 'Applied 2-up.')}>N-up 2</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'nUp', n: 4 }), 'Applied 4-up.')}>N-up 4</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => mergeRef.current?.click()}>Merge PDF…</DsButton>
                <input ref={mergeRef} className="visually-hidden" type="file" accept="application/pdf,.pdf" aria-label="Merge another PDF" onChange={(event) => { void mergePdf(event) }} />
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void (async () => {
                  if (!bytes) return
                  setBusy('editing')
                  try {
                    const part = await splitPdfPrefix(bytes, page)
                    downloadPdf(part, `${fileName.replace(/\.pdf$/i, '')}-split.pdf`)
                    setInfo(`Downloaded pages 1–${page}.`)
                  } catch (reason: unknown) {
                    setError(errorMessage(reason))
                  } finally {
                    setBusy(null)
                  }
                })()}>Download split 1–{page}</DsButton>
              </div>
            </div>
          )}

          {inspector === 'mark' && (
            <div className="ioc-panel ds-panel">
              <span className="ds-eyebrow">Text markup</span>
              <p className="ds-muted">Click text passages on the page to select them, then choose a style. Click again to deselect. Keyboard users can Tab to passages and press Space.</p>
              <p role="status" className="ds-muted">{selection.length} passage{selection.length === 1 ? '' : 's'} selected.</p>
              {!locked && !text.trim() && <p className="ds-muted">This page has no selectable text. Use Draw to mark an area instead.</p>}
              <div className="ds-row">
                {(['highlight', 'underline', 'strikeout'] as MarkupType[]).map((type) => (
                  <DsButton key={type} variant="outlined" className="workbench-button" disabled={locked || selection.length === 0} onClick={() => void runEdit(() => applyPdfMarkup(bytes!, page, type, selection), `Applied ${type} on page ${page}.`)}>{type}</DsButton>
                ))}
              </div>
              <span className="ds-eyebrow">Annotations</span>
              {annots.length === 0 ? <p className="ds-muted">No markup or notes yet.</p> : (
                <ol className="pdf-result-list">
                  {annots.map((annot) => (
                    <li key={`${annot.page}-${annot.objNum}`}>
                      <button type="button" className="workbench-button workbench-button--quiet" onClick={() => setPage(annot.page)}>Page {annot.page} · {annot.subtype}</button>
                      <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfAnnotDelete(bytes!, annot), `Deleted ${annot.subtype}.`)}>Delete</DsButton>
                    </li>
                  ))}
                </ol>
              )}
            </div>
          )}

          {inspector === 'draw' && (
            <div className="ioc-panel ds-panel">
              <span className="ds-eyebrow">Drawings and notes</span>
              <div className="ds-row">
                {(['rect', 'ellipse', 'line', 'arrow', 'ink'] as const).map(kind => <DsButton key={kind} variant="outlined" disabled={locked} aria-pressed={tool === kind} onClick={() => setTool(tool === kind ? null : kind)}>{kind === 'rect' ? 'Rectangle' : kind[0].toUpperCase() + kind.slice(1)}</DsButton>)}
              </div>
              <DsField label="Note contents"><DsInput value={noteText} onChange={(event) => setNoteText(event.target.value)} /></DsField>
              <DsButton variant="outlined" className="workbench-button" disabled={locked || !noteText.trim()} aria-pressed={tool === 'note'} onClick={() => setTool(tool === 'note' ? null : 'note')}>Place note</DsButton>
              {annots.filter((annot) => annot.subtype === 'note').map((annot) => (
                <DsButton key={`note-${annot.objNum}`} variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfNoteEdit(bytes!, annot, noteText), 'Edited note.')}>Save note {annot.objNum}</DsButton>
              ))}
            </div>
          )}

          {inspector === 'forms' && (
            <div className="ioc-panel ds-panel">
              <span className="ds-eyebrow">Fill form fields</span>
              <DsField label="Saved text appearance">
                <DsSelect aria-label="Saved text appearance" value={formAppearanceFont} disabled={locked || fieldBytes !== bytes} onChange={(event) => {
                  setFormAppearanceFont(event.target.value as TextAppearanceFont | 'viewer')
                  setFormNotice(null)
                }}>
                  <option value="viewer">Let the PDF viewer generate it</option>
                  <option value="Helvetica">Generate with Helvetica</option>
                  <option value="Times-Roman">Generate with Times Roman</option>
                  <option value="Courier">Generate with Courier</option>
                </DsSelect>
              </DsField>
              <p className="ds-muted">Choose a font to save fresh appearances for supported single-line text fields using printable ASCII. This replaces the text font; it does not preserve the original typography. Long text may clip in a fixed-size field. Unsupported text fields are skipped. Other field types may still depend on the PDF viewer.</p>
              {formNotice && <p role="status" className="ds-muted" aria-live="polite">{formNotice}</p>}
              {unsavedFormNames.length > 0 && <p role="status" className="ds-muted">Unsaved drafts: {unsavedFormNames.join(', ')}. These values are not in the downloaded PDF. Change the input or appearance mode and apply again.</p>}
              {fields.length === 0 ? <p className="ds-muted">No form fields in this file.</p> : fields.map((field, index) => (
                <DsField key={field.name} label={field.name}>
                  {field.kind === 'checkbox' ? (
                    <input type="checkbox" disabled={locked || fieldBytes !== bytes} checked={Boolean(field.checked)} onChange={(event) => setFields((current) => current.map((item, itemIndex) => itemIndex === index ? { ...item, checked: event.target.checked } : item))} />
                  ) : (
                    <DsInput disabled={locked || fieldBytes !== bytes} value={field.value ?? ''} onChange={(event) => setFields((current) => current.map((item, itemIndex) => itemIndex === index ? { ...item, value: event.target.value } : item))} />
                  )}
                </DsField>
              ))}
              <DsButton variant="outlined" className="workbench-button" disabled={locked || fieldBytes !== bytes || fields.length === 0} onClick={() => void applyFormDrafts()}>Apply form values</DsButton>
            </div>
          )}

          {inspector === 'stamp' && (
            <div className="ioc-panel ds-panel">
              <span className="ds-eyebrow">Stamps</span>
              <p className="ds-muted">These examples place a small sample image in the lower-left corner. A visual signature is an image, not a digital signature.</p>
              <div className="ds-stack">
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfStamp(bytes!, page, false), 'Applied an image stamp.')}>Apply sample image stamp</DsButton>
                <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfStamp(bytes!, page, true), 'Applied a visual signature stamp.')}>Apply sample visual signature</DsButton>
              </div>
            </div>
          )}

          {inspector === 'host' && (
            <>
              <DsCallout tone="note" title="Node host idle">Available during npm run dev. Static Pages builds stay browser-only.</DsCallout>
              <div className="ioc-panel ds-panel">
                <span className="ds-eyebrow">Node PDFium host</span>
                <p className="ds-muted">Text, image, OCR raster, and fail-closed redaction. Served by the Vite plugin at {PDF_NODE_PREFIX} during npm run dev.</p>
                <DsField label="Find"><DsInput value={hostOld} onChange={(event) => setHostOld(event.target.value)} /></DsField>
                <DsField label="Replace"><DsInput value={hostNew} onChange={(event) => setHostNew(event.target.value)} /></DsField>
                <div className="ds-stack">
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('text-edit', { oldText: hostOld, newText: hostNew, rect: current ? [0, 0, current.width, current.height] : [0, 0, 612, 792] })}>Surgical text edit</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('text-insert', { text: hostNew })}>Insert text</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('images/list')}>List images</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('images/insert')}>Insert image</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('images/transform')}>Rotate first image</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('images/delete')}>Delete first image</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('ocr/render')}>Rasterize page PNG</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('ocr/layer', { text: hostNew })}>OCR invisible layer</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('redact/text', { rect: current ? [72, current.height - 140, 360, current.height - 60] : [72, 650, 360, 730] })}>Redact text region</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('redact/image')}>Redact first image</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('redact/annotation')}>Redact first annotation</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('redact/path')}>Redact first path</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('redact/xobject')}>Redact first XObject</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={locked} onClick={() => runHost('redact/form')}>Redact first form field</DsButton>
                </div>
                <pre className="pdf-text">{hostNote}</pre>
              </div>
            </>
          )}
          </div>
        </aside>
      </div>
    </div>
  )
}
