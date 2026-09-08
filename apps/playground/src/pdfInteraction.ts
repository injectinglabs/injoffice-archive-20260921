export type PdfPoint = [number, number]
export type PdfRect = [number, number, number, number]

export function boundsOf(points: PdfPoint[]): PdfRect {
  return [Math.min(...points.map(p => p[0])), Math.min(...points.map(p => p[1])), Math.max(...points.map(p => p[0])), Math.max(...points.map(p => p[1]))]
}

/** Map client positions through the PDF.js viewport, including crop and rotation. */
export function pointerToPdf(client: PdfPoint, bounds: { left: number; top: number; width: number; height: number }, viewport: { width: number; height: number; convertToPdfPoint(x: number, y: number): number[] }): PdfPoint {
  const x = Math.max(0, Math.min(1, (client[0] - bounds.left) / bounds.width)) * viewport.width
  const y = Math.max(0, Math.min(1, (client[1] - bounds.top) / bounds.height)) * viewport.height
  const point = viewport.convertToPdfPoint(x, y)
  return [point[0]!, point[1]!]
}

export function rangesOverlap(start: number, length: number, match: number, queryLength: number): boolean {
  return start < match + queryLength && start + length > match
}

export type PdfSnapshot = { bytes: Uint8Array; page: number; edited: boolean }
export type PdfHistory = { past: PdfSnapshot[]; future: PdfSnapshot[] }
export function recordPdfEdit(history: PdfHistory, current: PdfSnapshot): PdfHistory {
  // Cap retained documents: uploaded PDFs can be large.
  return { past: [...history.past, current].slice(-20), future: [] }
}
export function travelPdfHistory(history: PdfHistory, current: PdfSnapshot, direction: 'undo' | 'redo'): { history: PdfHistory; snapshot: PdfSnapshot } | null {
  const source = direction === 'undo' ? history.past : history.future
  const snapshot = source.at(-1)
  if (!snapshot) return null
  return { snapshot, history: direction === 'undo'
    ? { past: source.slice(0, -1), future: [...history.future, current] }
    : { past: [...history.past, current], future: source.slice(0, -1) } }
}
