import { useEffect, useRef, useState, type ChangeEvent, type FormEvent } from 'react'
import pdfWorkerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import type { MarkupType } from '../../../../packages/pdf/src/annotate/types'
import type { PdfDocumentInfo } from '../../../../packages/pdf/src/types'
import {
  configurePdfWorker,
  PdfViewerDocument,
  renderPageToCanvas,
  type SearchMatch,
} from '../../../../packages/pdf/src/viewer'
import { PDF_NODE_PREFIX } from '../pdfHostRoutes'
import { createPdfDemoFixture, PDF_DEMO_FILE_NAME } from '../pdfDemoFixture'
import {
  applyPdfAnnotDelete,
  applyPdfDrawing,
  applyPdfFormValues,
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

configurePdfWorker(pdfWorkerUrl)

type OutlineItem = { title: string; pageIndex: number | null }
type Inspector = 'inspect' | 'pages' | 'mark' | 'draw' | 'forms' | 'stamp' | 'host'

const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2]
const INSPECTORS: { id: Inspector; label: string }[] = [
  { id: 'inspect', label: 'Inspect' },
  { id: 'pages', label: 'Pages' },
  { id: 'mark', label: 'Markup' },
  { id: 'draw', label: 'Draw' },
  { id: 'forms', label: 'Forms' },
  { id: 'stamp', label: 'Stamp' },
  { id: 'host', label: 'Node host' },
]

export function clampPdfPage(page: number, pageCount: number): number {
  if (pageCount < 1) return 1
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
  const [bytes, setBytes] = useState<Uint8Array | null>(null)
  const [fileName, setFileName] = useState(PDF_DEMO_FILE_NAME)
  const [edited, setEdited] = useState(false)
  const [viewer, setViewer] = useState<PdfViewerDocument | null>(null)
  const [page, setPage] = useState(1)
  const [pageCount, setPageCount] = useState(0)
  const [zoom, setZoom] = useState(1)
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
  const [noteText, setNoteText] = useState('Shared note')
  const [hostOld, setHostOld] = useState('InjOffice')
  const [hostNew, setHostNew] = useState('InjOffice PDF')
  const [hostNote, setHostNote] = useState('Node host idle. Available during npm run dev.')
  const [inspector, setInspector] = useState<Inspector>('pages')

  const applyBytes = async (next: Uint8Array, message: string) => {
    setBytes(next)
    setEdited(true)
    setInfo(message)
  }

  const runEdit = async (work: () => Promise<Uint8Array>, message: string) => {
    if (!bytes || busy) return
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
        const loaded = await PdfViewerDocument.load(bytes)
        ownedViewer = loaded
        if (cancelled) return
        viewerRef.current = loaded
        setPageCount(loaded.pageCount)
        setPage((current) => clampPdfPage(current, loaded.pageCount))
        setViewer(loaded)
        setBusy(null)
        setInfo(`${fileName} · ${loaded.pageCount} page${loaded.pageCount === 1 ? '' : 's'}`)
        const [nextOutline, nextGeometry, nextAnnots, nextFields] = await Promise.all([
          loaded.getOutline().catch(() => [] as OutlineItem[]),
          pdfGeometry(bytes),
          listPdfAnnots(bytes),
          listPdfFormFields(bytes),
        ])
        if (!cancelled && viewerRef.current === loaded) {
          setOutline(nextOutline)
          setGeometry(nextGeometry)
          setAnnots(nextAnnots)
          setFields(nextFields)
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
    let cancelled = false
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
        await renderPageToCanvas(viewer, selectedPage, canvas, zoom)
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
    return () => { cancelled = true }
  }, [fileName, page, pageCount, viewer, zoom])

  const openPdf = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0]
    event.currentTarget.value = ''
    if (!file) return
    setError(null)
    try {
      const nextBytes = new Uint8Array(await file.arrayBuffer())
      if (nextBytes.length === 0) throw new Error('the selected file is empty')
      setFileName(file.name || 'document.pdf')
      setEdited(false)
      setPage(1)
      setBytes(nextBytes)
    } catch (reason: unknown) {
      setError(`Could not read the selected PDF: ${errorMessage(reason)}`)
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
    setSearching(true)
    setError(null)
    try {
      const nextMatches = await activeViewer.search(query.trim())
      if (viewerRef.current !== activeViewer) return
      setMatches(nextMatches)
      setSearched(true)
      setSearching(false)
    } catch (reason: unknown) {
      if (viewerRef.current !== activeViewer) return
      setError(`Could not search the PDF: ${errorMessage(reason)}`)
      setSearching(false)
    }
  }

  const ready = viewer !== null && pageCount > 0
  const locked = !ready || busy !== null
  const current = geometry?.pages[page - 1]

  return (
    <div className="platen-fill workbench-surface" data-demo-surface="pdf">
      <div className="toolstrip pdf-toolbar workbench-toolbar" role="group" aria-label="PDF actions">
        <button type="button" className="workbench-button workbench-button--primary" onClick={() => uploadRef.current?.click()}>Open PDF</button>
        <input ref={uploadRef} className="visually-hidden" type="file" accept="application/pdf,.pdf" aria-label="Open a PDF file" onChange={(event) => { void openPdf(event) }} />
        <button type="button" className="workbench-button" disabled={!ready || busy !== null || page <= 1} onClick={() => setPage((currentPage) => clampPdfPage(currentPage - 1, pageCount))}>Previous</button>
        <label className="pdf-page-input">
          Page
          <span className="pdf-page-input__control">
            <input type="number" min={1} max={Math.max(1, pageCount)} value={page} disabled={!ready || busy === 'loading'} onChange={(event) => setPage(clampPdfPage(Number(event.target.value), pageCount))} />
            <span>of {pageCount || '—'}</span>
          </span>
        </label>
        <button type="button" className="workbench-button" disabled={!ready || busy !== null || page >= pageCount} onClick={() => setPage((currentPage) => clampPdfPage(currentPage + 1, pageCount))}>Next</button>
        <label className="tool-field">
          <span>Zoom</span>
          <select aria-label="Zoom" value={zoom} disabled={!ready} onChange={(event) => setZoom(Number(event.target.value))}>{ZOOM_STEPS.map((step) => <option key={step} value={step}>{Math.round(step * 100)}%</option>)}</select>
        </label>
        <button type="button" className="workbench-button" disabled={!bytes || busy === 'loading'} onClick={() => downloadPdf()}>Download {edited ? 'edited PDF' : 'PDF'}</button>
      </div>
      <p className="native-status workbench-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : busy ?? 'ready'}>
        {busy === 'rendering' ? 'Rendering…' : busy === 'editing' ? 'Applying PDF operation…' : info}
      </p>
      {error && <p className="native-error pdf-error workbench-callout workbench-callout--error" role="alert">{error}</p>}
      <div className="split">
        <div className="split-main pages pdf-pages" aria-busy={busy === 'loading' || busy === 'rendering' || busy === 'editing'}>
          {!ready && !error && <p className="pdf-empty">Loading PDF…</p>}
          <canvas ref={canvasRef} className="pdf-canvas" aria-label={ready ? `Rendered page ${page} of ${pageCount}` : 'PDF page'} />
        </div>
        <aside className="split-side pdf-inspector workbench-inspector" aria-label="PDF operations">
          <div className="view-switcher" role="tablist" aria-label="PDF tool panels">
            <div className="tool-segment">
              {INSPECTORS.map((item) => (
                <button key={item.id} type="button" role="tab" aria-selected={inspector === item.id} aria-pressed={inspector === item.id} onClick={() => setInspector(item.id)}>{item.label}</button>
              ))}
            </div>
          </div>

          {inspector === 'inspect' && (
            <>
              <form className="ioc-panel pdf-search" onSubmit={(event) => { void search(event) }}>
                <strong>Search document</strong>
                <div className="pdf-search-row">
                  <input aria-label="Search PDF text" type="search" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Find extracted text" />
                  <button type="submit" className="workbench-button" disabled={!ready || searching || query.trim().length === 0}>{searching ? 'Finding…' : 'Find'}</button>
                </div>
                {searched && <small>{matches.length} match{matches.length === 1 ? '' : 'es'}</small>}
                {matches.length > 0 && <ol className="pdf-result-list">{matches.map((match, index) => <li key={`${match.pageIndex}-${match.offset}-${index}`}><button type="button" className="workbench-button workbench-button--quiet" onClick={() => setPage(match.pageIndex)}><span>Page {match.pageIndex}</span>{match.snippet || query}</button></li>)}</ol>}
              </form>
              <div className="ioc-panel"><strong>Geometry</strong><p>{geometry ? `${geometry.pageCount} pages` : '—'}</p>{current && <p>{Math.round(current.width)}×{Math.round(current.height)} pt · {current.rotation}°</p>}</div>
              <div className="ioc-panel pdf-outline"><strong>Document outline</strong>{outline.length === 0 ? <p>No outline entries.</p> : <ol className="pdf-result-list">{outline.map((item, index) => <li key={`${item.title}-${index}`}><button type="button" className="workbench-button workbench-button--quiet" disabled={item.pageIndex === null} onClick={() => item.pageIndex !== null && setPage(item.pageIndex)}>{item.title}{item.pageIndex === null ? '' : ` · ${item.pageIndex}`}</button></li>)}</ol>}</div>
              <div className="ioc-panel pdf-text"><strong>Page {page} extracted text</strong><p>{text || 'No extractable text on this page.'}</p></div>
            </>
          )}

          {inspector === 'pages' && (
            <div className="ioc-panel">
              <strong>Page operations</strong>
              <p>Browser-safe `applyPageOps`, merge, and split.</p>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'rotate', pages: [page], degrees: 90 }), `Rotated page ${page} by 90°.`)}>Rotate 90°</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'rotate', pages: [page], degrees: 180 }), `Rotated page ${page} by 180°.`)}>Rotate 180°</button>
              <button type="button" className="workbench-button" disabled={locked || pageCount < 2} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'delete', pages: [page] }), `Deleted page ${page}.`)}>Delete page</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'insertBlank', at: page + 1 }), `Inserted a blank page after ${page}.`)}>Insert blank after</button>
              <button type="button" className="workbench-button" disabled={locked || !movePageOrder(pageCount, page, -1)} onClick={() => { const order = movePageOrder(pageCount, page, -1); if (order) void runEdit(() => applyPdfPageOp(bytes!, { type: 'reorder', order }), 'Moved page earlier.') }}>Move earlier</button>
              <button type="button" className="workbench-button" disabled={locked || !movePageOrder(pageCount, page, 1)} onClick={() => { const order = movePageOrder(pageCount, page, 1); if (order) void runEdit(() => applyPdfPageOp(bytes!, { type: 'reorder', order }), 'Moved page later.') }}>Move later</button>
              <button type="button" className="workbench-button" disabled={locked || !current} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'crop', pages: [page], box: { x: 36, y: 36, width: Math.max(72, current!.width - 72), height: Math.max(72, current!.height - 72) } }), `Cropped page ${page}.`)}>Crop 36pt inset</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'resize', pages: [page], width: 612, height: 792, fit: 'contain' }), `Resized page ${page} to letter.`)}>Resize to letter</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'nUp', n: 2 }), 'Applied 2-up.')}>N-up 2</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfPageOp(bytes!, { type: 'nUp', n: 4 }), 'Applied 4-up.')}>N-up 4</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => mergeRef.current?.click()}>Merge PDF…</button>
              <input ref={mergeRef} className="visually-hidden" type="file" accept="application/pdf,.pdf" aria-label="Merge another PDF" onChange={(event) => { void mergePdf(event) }} />
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void (async () => {
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
              })()}>Download split 1–{page}</button>
            </div>
          )}

          {inspector === 'mark' && (
            <div className="ioc-panel">
              <strong>Text markup</strong>
              {(['highlight', 'underline', 'strikeout'] as MarkupType[]).map((type) => (
                <button key={type} type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfMarkup(bytes!, page, type), `Applied ${type} on page ${page}.`)}>{type}</button>
              ))}
              <strong>Annotations</strong>
              {annots.length === 0 ? <p>No markup or notes yet.</p> : (
                <ol className="pdf-result-list">
                  {annots.map((annot) => (
                    <li key={`${annot.page}-${annot.objNum}`}>
                      <button type="button" className="workbench-button workbench-button--quiet" onClick={() => setPage(annot.page)}>Page {annot.page} · {annot.subtype}</button>
                      <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfAnnotDelete(bytes!, annot), `Deleted ${annot.subtype}.`)}>Delete</button>
                    </li>
                  ))}
                </ol>
              )}
            </div>
          )}

          {inspector === 'draw' && (
            <div className="ioc-panel">
              <strong>Drawings and notes</strong>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfDrawing(bytes!, page, 'rect'), 'Drew a rectangle.')}>Rectangle</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfDrawing(bytes!, page, 'ellipse'), 'Drew an ellipse.')}>Ellipse</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfDrawing(bytes!, page, 'line'), 'Drew a line.')}>Line</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfDrawing(bytes!, page, 'arrow'), 'Drew an arrow.')}>Arrow</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfDrawing(bytes!, page, 'ink'), 'Drew an ink stroke.')}>Ink</button>
              <label>Note contents<input value={noteText} onChange={(event) => setNoteText(event.target.value)} /></label>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfNote(bytes!, page, noteText), 'Added a sticky note.')}>Add note</button>
              {annots.filter((annot) => annot.subtype === 'note').map((annot) => (
                <button key={`note-${annot.objNum}`} type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfNoteEdit(bytes!, annot, noteText), 'Edited note.')}>Save note {annot.objNum}</button>
              ))}
            </div>
          )}

          {inspector === 'forms' && (
            <div className="ioc-panel">
              <strong>AcroForm values</strong>
              {fields.length === 0 ? <p>No form fields in this file.</p> : fields.map((field, index) => (
                <label key={field.name}>
                  {field.name}
                  {field.kind === 'checkbox' ? (
                    <input type="checkbox" checked={Boolean(field.checked)} onChange={(event) => setFields((current) => current.map((item, itemIndex) => itemIndex === index ? { ...item, checked: event.target.checked } : item))} />
                  ) : (
                    <input value={field.value ?? ''} onChange={(event) => setFields((current) => current.map((item, itemIndex) => itemIndex === index ? { ...item, value: event.target.value } : item))} />
                  )}
                </label>
              ))}
              <button type="button" className="workbench-button" disabled={locked || fields.length === 0} onClick={() => void runEdit(() => applyPdfFormValues(bytes!, fields), 'Applied form values.')}>Apply form values</button>
            </div>
          )}

          {inspector === 'stamp' && (
            <div className="ioc-panel">
              <strong>Stamps</strong>
              <p>Uses the package stamp/signature helpers with a 1×1 PNG fixture.</p>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfStamp(bytes!, page, false), 'Applied an image stamp.')}>Image stamp</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => void runEdit(() => applyPdfStamp(bytes!, page, true), 'Applied a visual signature stamp.')}>Visual signature stamp</button>
            </div>
          )}

          {inspector === 'host' && (
            <div className="ioc-panel">
              <strong>Node PDFium host</strong>
              <p>Text, image, OCR raster, and fail-closed redaction. Served by the Vite plugin at {PDF_NODE_PREFIX} during <code>npm run dev</code>.</p>
              <label>Find<input value={hostOld} onChange={(event) => setHostOld(event.target.value)} /></label>
              <label>Replace<input value={hostNew} onChange={(event) => setHostNew(event.target.value)} /></label>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('text-edit', { oldText: hostOld, newText: hostNew, rect: current ? [0, 0, current.width, current.height] : [0, 0, 612, 792] })}>Surgical text edit</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('text-insert', { text: hostNew })}>Insert text</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('images/list')}>List images</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('images/insert')}>Insert image</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('images/transform')}>Rotate first image</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('images/delete')}>Delete first image</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('ocr/render')}>Rasterize page PNG</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('ocr/layer', { text: hostNew })}>OCR invisible layer</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('redact/text', { rect: current ? [72, current.height - 140, 360, current.height - 60] : [72, 650, 360, 730] })}>Redact text region</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('redact/image')}>Redact first image</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('redact/annotation')}>Redact first annotation</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('redact/path')}>Redact first path</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('redact/xobject')}>Redact first XObject</button>
              <button type="button" className="workbench-button" disabled={locked} onClick={() => runHost('redact/form')}>Redact first form field</button>
              <pre className="pdf-text">{hostNote}</pre>
            </div>
          )}
        </aside>
      </div>
    </div>
  )
}
