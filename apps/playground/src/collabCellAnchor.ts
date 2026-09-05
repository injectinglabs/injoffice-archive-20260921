/**
 * Browser-safe geometry for React-owned overlays above a Univer sheet.
 *
 * Univer's `FRange.getCellRect()` returns sheet-skeleton coordinates. They
 * include the row/column headers, but intentionally do not include zoom,
 * scrolling, responsive canvas scaling, or the canvas' page position. This
 * module applies those transforms without registering a Univer popup.
 */

export interface RectLike {
  left: number
  top: number
  right: number
  bottom: number
  width: number
  height: number
}

export interface SheetViewportScroll {
  viewportScrollX: number
  viewportScrollY: number
}

export interface SheetFreeze {
  startColumn: number
  startRow: number
}

export interface SheetRangeLike {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

export interface CellAnchor {
  /** Coordinates relative to the element that owns the overlay. */
  left: number
  top: number
  right: number
  bottom: number
  width: number
  height: number
}

export interface CellAnchorTransform {
  cellRect: RectLike
  canvasRect: RectLike
  overlayRect: Pick<RectLike, 'left' | 'top'>
  canvasCssWidth?: number
  canvasCssHeight?: number
  zoom: number
  viewportScrollX: number
  viewportScrollY: number
}

export interface UniverRangeLike {
  getCellRect(): RectLike
}

export interface UniverWorksheetLike {
  getRange(row: number, column: number): UniverRangeLike
  getZoom(): number
  getFreeze?(): SheetFreeze
  getVisibleRangesOfAllViewports?(): Map<unknown, SheetRangeLike> | null
}

export interface UniverCanvasLike {
  getBoundingClientRect(): RectLike
  style: { width: string; height: string }
}

export interface UniverEditorRootLike {
  querySelector(selector: string): UniverCanvasLike | null
}

export interface OverlayRootLike {
  getBoundingClientRect(): RectLike
}

function finiteOr(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback
}

function positiveOr(value: number | undefined, fallback: number): number {
  return value !== undefined && Number.isFinite(value) && value > 0 ? value : fallback
}

/** Parse a CSS pixel length without treating an empty/invalid value as zero. */
export function cssPixels(value: string | undefined): number | undefined {
  if (!value) return undefined
  const parsed = Number.parseFloat(value)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined
}

/**
 * Apply the same transform as Univer 0.25's
 * `SheetCanvasPopManagerService._calcCellPositionByCell`, then express it
 * relative to our React overlay root rather than the browser viewport.
 */
export function mapCellRectToOverlay(transform: CellAnchorTransform): CellAnchor | null {
  const {
    cellRect,
    canvasRect,
    overlayRect,
    viewportScrollX,
    viewportScrollY,
  } = transform
  const zoom = positiveOr(transform.zoom, 1)
  const scaleX = canvasRect.width / positiveOr(transform.canvasCssWidth, canvasRect.width)
  const scaleY = canvasRect.height / positiveOr(transform.canvasCssHeight, canvasRect.height)

  const pageLeft = canvasRect.left + (cellRect.left - finiteOr(viewportScrollX, 0)) * scaleX * zoom
  const pageRight = canvasRect.left + (cellRect.right - finiteOr(viewportScrollX, 0)) * scaleX * zoom
  const pageTop = canvasRect.top + (cellRect.top - finiteOr(viewportScrollY, 0)) * scaleY * zoom
  const pageBottom = canvasRect.top + (cellRect.bottom - finiteOr(viewportScrollY, 0)) * scaleY * zoom

  if (![pageLeft, pageRight, pageTop, pageBottom].every(Number.isFinite)) return null

  const left = pageLeft - overlayRect.left
  const right = pageRight - overlayRect.left
  const top = pageTop - overlayRect.top
  const bottom = pageBottom - overlayRect.top
  return {
    left,
    right,
    top,
    bottom,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  }
}

/**
 * Frozen panes do not scroll on their frozen axis. The main viewport uses
 * both scroll values; top/left panes use only the remaining axis.
 */
export function scrollForCell(
  scroll: SheetViewportScroll,
  freeze: SheetFreeze | null | undefined,
  row: number,
  column: number,
): SheetViewportScroll {
  if (!freeze) return scroll
  return {
    viewportScrollX: column < freeze.startColumn ? 0 : scroll.viewportScrollX,
    viewportScrollY: row < freeze.startRow ? 0 : scroll.viewportScrollY,
  }
}

export function cellIsVisible(row: number, column: number, ranges: Iterable<SheetRangeLike> | null | undefined): boolean {
  if (!ranges) return true
  for (const range of ranges) {
    if (row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn) return true
  }
  return false
}

/**
 * Measure one cell in one editor instance. Querying the render canvas below
 * `editorRoot` is important: a page may mount several independent Univer
 * instances and a document-wide canvas query can anchor to the wrong sheet.
 */
export function measureUniverCellAnchor(options: {
  editorRoot: UniverEditorRootLike
  overlayRoot: OverlayRootLike
  worksheet: UniverWorksheetLike
  row: number
  column: number
  scroll: SheetViewportScroll
}): CellAnchor | null {
  const { editorRoot, overlayRoot, worksheet, row, column } = options
  const visible = worksheet.getVisibleRangesOfAllViewports?.()
  if (visible && !cellIsVisible(row, column, visible.values())) return null

  const canvas = editorRoot.querySelector('canvas[data-u-comp="render-canvas"]')
  if (!canvas) return null

  const scroll = scrollForCell(options.scroll, worksheet.getFreeze?.(), row, column)
  return mapCellRectToOverlay({
    cellRect: worksheet.getRange(row, column).getCellRect(),
    canvasRect: canvas.getBoundingClientRect(),
    overlayRect: overlayRoot.getBoundingClientRect(),
    canvasCssWidth: cssPixels(canvas.style.width),
    canvasCssHeight: cssPixels(canvas.style.height),
    zoom: worksheet.getZoom(),
    viewportScrollX: scroll.viewportScrollX,
    viewportScrollY: scroll.viewportScrollY,
  })
}
