import { useEffect, useMemo, useRef, useState } from 'react'
import { renderPageToCanvas, type PdfViewerDocument } from '../../../../packages/pdf/src/viewer'
import { PdfTextLayer } from './PdfTextLayer'
import { createPdfRenderQueue, pdfPageWindow } from './pdfRenderQueue'
import './pdf-continuous.css'

type Queue = ReturnType<typeof createPdfRenderQueue>

function ReadPage({ viewer, page, zoom, queue, thumbnail = false, onSize }: {
  viewer: PdfViewerDocument; page: number; zoom: number; queue: Queue; thumbnail?: boolean
  onSize?: (width: number, height: number) => void
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const [state, setState] = useState('loading')
  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const controller = new AbortController()
    setState('loading')
    void queue(controller.signal, async () => {
      const size = await viewer.getPageSize(page)
      controller.signal.throwIfAborted()
      const scale = thumbnail ? Math.min(100 / size.width, 130 / size.height) : zoom
      if (!thumbnail) onSize?.(size.width * scale, size.height * scale)
      await renderPageToCanvas(viewer, page, canvas, scale, thumbnail ? 1 : undefined, {
        signal: controller.signal, maxPixels: thumbnail ? 20_000 : 4_000_000,
      })
      if (!controller.signal.aborted) setState('ready')
    }).catch(() => { if (!controller.signal.aborted) setState('error') })
    return () => { controller.abort(); canvas.width = 0; canvas.height = 0 }
  }, [viewer, page, zoom, queue, thumbnail, onSize])
  return <div className={thumbnail ? 'pdf-thumbnail-image' : 'pdf-page-stage'} data-render-state={state}>
    <canvas ref={canvasRef} aria-label={thumbnail ? `Thumbnail of page ${page}` : `Rendered page ${page}`} />
    {!thumbnail && state === 'ready' && <PdfTextLayer viewer={viewer} page={page} zoom={zoom} disabled={false} />}
    {state === 'error' && <span role="status">Page {page} preview unavailable.</span>}
  </div>
}

/** Only the selected page and its neighbours own full-size canvases. The rail
 * retains five tiny previews. Offscreen pages remain lightweight placeholders. */
export function PdfContinuousView({ viewer, page, zoom, onPageChange }: {
  viewer: PdfViewerDocument; page: number; zoom: number; onPageChange: (page: number) => void
}) {
  const scrollRef = useRef<HTMLDivElement>(null)
  const railRef = useRef<HTMLElement>(null)
  const observedPage = useRef<number | null>(null)
  const queue = useMemo(createPdfRenderQueue, [viewer])
  const [sizes, setSizes] = useState<Record<number, { width: number; height: number }>>({})
  const callbacks = useMemo(() => new Map<number, (width: number, height: number) => void>(), [viewer, zoom])
  const active = new Set(pdfPageWindow(page, viewer.pageCount))
  const thumbnails = new Set(pdfPageWindow(page, viewer.pageCount, 2))
  const pages = Array.from({ length: viewer.pageCount }, (_, index) => index + 1)
  useEffect(() => { setSizes({}) }, [viewer, zoom])
  useEffect(() => {
    const rail = railRef.current
    if (!rail) return
    const observer = new ResizeObserver(() => {
      const button = rail.querySelector<HTMLElement>('[aria-current="page"]')
      if (!button) return
      const bounds = rail.getBoundingClientRect()
      const item = button.getBoundingClientRect()
      if (item.top < bounds.top) rail.scrollTop += item.top - bounds.top
      else if (item.bottom > bounds.bottom) rail.scrollTop += item.bottom - bounds.bottom
    })
    observer.observe(rail)
    return () => observer.disconnect()
  }, [viewer])
  useEffect(() => {
    const rail = railRef.current
    const button = rail?.querySelector<HTMLElement>('[aria-current="page"]')
    if (rail && button) {
      const bounds = rail.getBoundingClientRect()
      const item = button.getBoundingClientRect()
      if (item.top < bounds.top) rail.scrollTop += item.top - bounds.top
      else if (item.bottom > bounds.bottom) rail.scrollTop += item.bottom - bounds.bottom
    }
    if (observedPage.current === page) { observedPage.current = null; return }
    const root = scrollRef.current
    const target = root?.querySelector<HTMLElement>(`[data-read-page="${page}"]`)
    if (root && target) root.scrollTop += target.getBoundingClientRect().top - root.getBoundingClientRect().top - 16
  }, [page, viewer, zoom])
  const measure = (index: number) => {
    if (!callbacks.has(index)) callbacks.set(index, (width, height) => setSizes(previous => previous[index]?.width === width && previous[index]?.height === height ? previous : { ...previous, [index]: { width, height } }))
    return callbacks.get(index)!
  }
  const trackScroll = () => {
    const root = scrollRef.current
    if (!root) return
    const top = root.getBoundingClientRect().top + 24
    // Page slots are ordered vertically. Binary search avoids measuring every
    // placeholder on each scroll event in very long PDFs.
    let low = 0
    let high = root.children.length - 1
    while (low < high) {
      const middle = Math.floor((low + high) / 2)
      if (root.children[middle]!.getBoundingClientRect().bottom < top) low = middle + 1
      else high = middle
    }
    const closest = low + 1
    if (closest !== page) { observedPage.current = closest; onPageChange(closest) }
  }
  return <div className="pdf-reader">
    <nav ref={railRef} className="pdf-thumbnail-rail" aria-label="PDF page thumbnails">
      {pages.map(index => <button key={index} type="button" aria-label={`Go to page ${index}`} aria-current={index === page ? 'page' : undefined} onClick={() => {
        observedPage.current = null
        if (index === page) {
          const root = scrollRef.current
          const target = root?.querySelector<HTMLElement>(`[data-read-page="${index}"]`)
          if (root && target) root.scrollTop += target.getBoundingClientRect().top - root.getBoundingClientRect().top - 16
        } else onPageChange(index)
      }}>
        {thumbnails.has(index) ? <ReadPage viewer={viewer} page={index} zoom={1} queue={queue} thumbnail /> : <span className="pdf-thumbnail-placeholder" aria-hidden="true" />}
        <span>Page {index}</span>
      </button>)}
    </nav>
    <div ref={scrollRef} className="pdf-continuous-scroll" role="region" aria-label="Continuous PDF pages" tabIndex={0} onScroll={trackScroll}>
      {pages.map(index => <article key={index} data-read-page={index} aria-label={`Page ${index}`} style={{ width: sizes[index]?.width ?? 612 * zoom, minHeight: sizes[index]?.height ?? 792 * zoom }}>
        {active.has(index) ? <ReadPage viewer={viewer} page={index} zoom={zoom} queue={queue} onSize={measure(index)} /> : <span className="pdf-page-placeholder">Page {index}</span>}
      </article>)}
    </div>
  </div>
}
