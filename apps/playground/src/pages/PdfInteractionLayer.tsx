import { useEffect, useState, type PointerEvent } from 'react'
import type { PageViewport } from 'pdfjs-dist'
import type { PdfViewerDocument, SearchMatch } from '../../../../packages/pdf/src/viewer'
import { boundsOf, pointerToPdf, rangesOverlap, type PdfPoint } from '../pdfInteraction'
import './pdf-interaction.css'

export type PdfPlacementTool = 'rect' | 'ellipse' | 'line' | 'arrow' | 'ink' | 'note' | null
type TextRun = { text: string; offset: number; quad: number[]; box: number[] }

export function PdfInteractionLayer({ viewer, page, zoom, locked, selectText, tool, matches, queryLength, onSelection, onPlace }: {
  viewer: PdfViewerDocument; page: number; zoom: number; locked: boolean; selectText: boolean; tool: PdfPlacementTool
  matches: SearchMatch[]; queryLength: number; onSelection: (quads: number[][]) => void
  onPlace: (tool: NonNullable<PdfPlacementTool>, points: PdfPoint[]) => void
}) {
  const [layout, setLayout] = useState<{ viewport: PageViewport; runs: TextRun[] } | null>(null)
  const [selected, setSelected] = useState<number[]>([])
  const [stroke, setStroke] = useState<PdfPoint[]>([])
  const [keyboardPoint, setKeyboardPoint] = useState<PdfPoint | null>(null)
  useEffect(() => {
    let cancelled = false
    setLayout(null)
    setSelected([])
    setStroke([])
    setKeyboardPoint(null)
    void viewer.getPage(page).then(async proxy => {
      const viewport = proxy.getViewport({ scale: zoom })
      const content = await proxy.getTextContent()
      let offset = 0
      const runs: TextRun[] = []
      for (const item of content.items) {
        const start = offset
        offset += ('str' in item ? item.str.length : 0) + 1
        if (!('str' in item) || !item.str.trim()) continue
        const [a, b, , , x, y] = item.transform as [number, number, number, number, number, number]
        const length = Math.hypot(a, b) || 1
        const ux = a / length, uy = b / length
        const height = item.height || length
        const bottomLeft: PdfPoint = [x + uy * height * 0.2, y - ux * height * 0.2]
        const topLeft: PdfPoint = [bottomLeft[0] - uy * height, bottomLeft[1] + ux * height]
        const bottomRight: PdfPoint = [bottomLeft[0] + ux * item.width, bottomLeft[1] + uy * item.width]
        const topRight: PdfPoint = [topLeft[0] + ux * item.width, topLeft[1] + uy * item.width]
        const corners = [topLeft, topRight, bottomLeft, bottomRight]
        const box = boundsOf(corners.map(p => viewport.convertToViewportPoint(...p) as PdfPoint))
        runs.push({ text: item.str, offset: start, quad: corners.flat(), box })
      }
      if (!cancelled) setLayout({ viewport, runs })
    }).catch(() => { if (!cancelled) setLayout(null) })
    return () => { cancelled = true }
  }, [viewer, page, zoom])

  if (!layout || locked) return null
  const { viewport, runs } = layout
  const point = (event: PointerEvent<SVGSVGElement>): PdfPoint => pointerToPdf([event.clientX, event.clientY], event.currentTarget.getBoundingClientRect(), viewport)
  const screenStroke = stroke.map(p => viewport.convertToViewportPoint(...p) as PdfPoint)
  const box = screenStroke.length ? boundsOf(screenStroke) : null
  return <div className="pdf-interaction" style={{ width: viewport.width, height: viewport.height }}>
    {runs.map((run, index) => {
      const found = matches.some(match => match.pageIndex === page && rangesOverlap(run.offset, run.text.length, match.offset, queryLength))
      const style = { left: run.box[0], top: run.box[1], width: run.box[2]! - run.box[0]!, height: run.box[3]! - run.box[1]! }
      return selectText
        ? <button key={index} type="button" className="pdf-text-target" style={style} aria-label={`Select text: ${run.text}`} aria-pressed={selected.includes(index)} data-match={found || undefined} onClick={() => {
          const next = selected.includes(index) ? selected.filter(i => i !== index) : [...selected, index]
          setSelected(next)
          onSelection(next.map(i => runs[i]!.quad))
        }} />
        : found ? <span key={index} className="pdf-search-highlight" style={style} aria-label={`Search match in: ${run.text}`} /> : null
    })}
    {tool && <svg className="pdf-drawing-surface" width={viewport.width} height={viewport.height} role="img" tabIndex={0} aria-label={`${tool} placement area. Drag on the page, or use arrow keys and Enter to set start and end points. Escape cancels.`}
      onFocus={() => setKeyboardPoint(previous => previous ?? [viewport.width / 2, viewport.height / 2])}
      onKeyDown={event => {
        const cursor = keyboardPoint ?? [viewport.width / 2, viewport.height / 2]
        const step = event.shiftKey ? 20 : 5
        const delta = event.key === 'ArrowLeft' ? [-step, 0] : event.key === 'ArrowRight' ? [step, 0] : event.key === 'ArrowUp' ? [0, -step] : event.key === 'ArrowDown' ? [0, step] : null
        if (delta) {
          event.preventDefault()
          const next: PdfPoint = [Math.max(0, Math.min(viewport.width, cursor[0]! + delta[0]!)), Math.max(0, Math.min(viewport.height, cursor[1]! + delta[1]!))]
          setKeyboardPoint(next)
          if (stroke.length) {
            const pdfPoint = viewport.convertToPdfPoint(...next) as PdfPoint
            setStroke(previous => tool === 'ink' ? [...previous, pdfPoint] : [previous[0]!, pdfPoint])
          }
        } else if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault()
          const pdfPoint = viewport.convertToPdfPoint(cursor[0]!, cursor[1]!) as PdfPoint
          if (tool === 'note') onPlace(tool, [pdfPoint])
          else if (!stroke.length) setStroke([pdfPoint])
          else {
            const points = [...stroke, pdfPoint]
            if (Math.hypot(pdfPoint[0] - points[0]![0], pdfPoint[1] - points[0]![1]) > 2 || (tool === 'ink' && points.length > 3)) { setStroke([]); onPlace(tool, points) }
          }
        }
      }}
      onPointerDown={event => { if (event.button !== 0) return; event.currentTarget.setPointerCapture(event.pointerId); setStroke([point(event)]) }}
      onPointerMove={event => { if (!event.currentTarget.hasPointerCapture(event.pointerId)) return; const next = point(event); setStroke(previous => tool === 'ink' ? [...previous, next] : [previous[0]!, next]) }}
      onPointerCancel={() => setStroke([])}
      onPointerUp={event => {
        if (!event.currentTarget.hasPointerCapture(event.pointerId) || !stroke.length) return
        event.currentTarget.releasePointerCapture(event.pointerId)
        const points = [...stroke, point(event)]
        setStroke([])
        if (tool === 'note' || Math.hypot(points.at(-1)![0] - points[0]![0], points.at(-1)![1] - points[0]![1]) > 2 || (tool === 'ink' && points.length > 3)) onPlace(tool, points)
      }}>
      {box && (tool === 'rect' ? <rect x={box[0]} y={box[1]} width={box[2] - box[0]} height={box[3] - box[1]} />
        : tool === 'ellipse' ? <ellipse cx={(box[0] + box[2]) / 2} cy={(box[1] + box[3]) / 2} rx={(box[2] - box[0]) / 2} ry={(box[3] - box[1]) / 2} />
        : <polyline points={screenStroke.map(p => p.join(',')).join(' ')} />)}
      {keyboardPoint && <path d={`M ${keyboardPoint[0] - 6} ${keyboardPoint[1]} h 12 M ${keyboardPoint[0]} ${keyboardPoint[1] - 6} v 12`} />}
    </svg>}
  </div>
}
