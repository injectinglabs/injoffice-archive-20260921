import { useEffect, useRef, useState } from 'react'
import { TextLayer } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PdfViewerDocument } from '../../../../packages/pdf/src/viewer'
import './pdf-text-layer.css'

/** PDF.js owns text positioning, including page rotation and font transforms.
 * Annotation tools keep their own PDF-coordinate targets above this read layer.
 */
export function PdfTextLayer({ viewer, page, zoom, disabled }: {
  viewer: PdfViewerDocument; page: number; zoom: number; disabled: boolean
}) {
  const containerRef = useRef<HTMLDivElement>(null)
  const [failed, setFailed] = useState(false)
  useEffect(() => {
    const container = containerRef.current
    if (!container) return
    let cancelled = false
    let layer: TextLayer | null = null
    setFailed(false)
    container.replaceChildren()
    void (async () => {
      const proxy = await viewer.getPage(page)
      if (cancelled) return
      const viewport = proxy.getViewport({ scale: zoom })
      container.style.setProperty('--total-scale-factor', String(viewport.scale * viewport.userUnit))
      layer = new TextLayer({
        textContentSource: proxy.streamTextContent(),
        container,
        viewport,
      })
      await layer.render()
    })().catch(() => {
      if (!cancelled) { container.replaceChildren(); setFailed(true) }
    })
    return () => {
      cancelled = true
      layer?.cancel()
      container.replaceChildren()
    }
  }, [viewer, page, zoom])

  return <>
    <div ref={containerRef} className="pdf-selectable-text" data-disabled={disabled || undefined} aria-hidden={disabled || undefined} aria-label={`Selectable text for page ${page}`} />
    {failed && <span className="visually-hidden" role="status">Selectable text unavailable. The rendered page remains available.</span>}
  </>
}
