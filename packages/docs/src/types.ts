export type PageSize = 'letter' | 'a4'

/** CSS px per inch at 100% zoom (96 dpi — the standard CSS reference pixel). */
export const CSS_PX_PER_INCH = 96

/** Page dimensions in CSS px, portrait, at 100% zoom. */
export const PAGE_DIMENSIONS_PX: Record<PageSize, { width: number; height: number }> = {
  letter: { width: 8.5 * CSS_PX_PER_INCH, height: 11 * CSS_PX_PER_INCH },
  a4: {
    width: (210 / 25.4) * CSS_PX_PER_INCH,
    height: (297 / 25.4) * CSS_PX_PER_INCH,
  },
}

export interface Margins {
  top: number
  bottom: number
  left: number
  right: number
}

/** Word/Excel-style default margins (1 inch all around), in CSS px. */
export const DEFAULT_MARGINS_PX: Margins = {
  top: CSS_PX_PER_INCH,
  bottom: CSS_PX_PER_INCH,
  left: CSS_PX_PER_INCH,
  right: CSS_PX_PER_INCH,
}

export interface PaginationOptions {
  pageSize: PageSize
  margins: Margins
  /** px reserved for the header band (0 = no header). */
  headerHeight: number
  /** px reserved for the footer band (0 = no footer). */
  footerHeight: number
}

export const DEFAULT_PAGINATION_OPTIONS: PaginationOptions = {
  pageSize: 'letter',
  margins: { ...DEFAULT_MARGINS_PX },
  headerHeight: 0,
  footerHeight: 0,
}

/** One line's extent within its owning block, used as a split opportunity. */
export interface LineBox {
  /** the line's top, in px, relative to the owning block's own top (0 = block's first line). */
  offset: number
  /** the line's height in px. */
  height: number
}

/** One top-level flow block, measured from the live DOM by the thin shell. */
export interface BlockBox {
  /** stable key: the block's ProseMirror position at measurement time. */
  id: string
  /** block height in px, border-box. */
  height: number
  /** a manual page-break node (docPageBreak) sits immediately before this
   *  block, or the block itself carries pageBreakBefore: force a new page. */
  breakBefore?: boolean
  /** Measured line extents, in ascending document order. */
  lineBoxes?: LineBox[]
}

/** One page's worth of blocks, by index into the original blocks array. */
export interface PageSlice {
  /** Inclusive start and exclusive end indices into the input block array. */
  start: number
  end: number
  /** Relative Y offset at which a continued first block begins. */
  startOffset?: number
  /** Relative Y offset at which a split final block ends. */
  endOffset?: number
}

export interface PageLayout {
  pages: PageSlice[]
  /** content-area height available per page (px), after margins/header/footer. */
  contentHeight: number
}

function pageDimension(opts: PaginationOptions) {
  const dimensions = PAGE_DIMENSIONS_PX[opts.pageSize]
  if (!dimensions) throw new RangeError(`unsupported page size: ${String(opts.pageSize)}`)
  return dimensions
}

/** content height available inside one page, after margins + header/footer bands. */
export function contentHeightPx(opts: PaginationOptions): number {
  return pageDimension(opts).height
    - opts.margins.top
    - opts.margins.bottom
    - opts.headerHeight
    - opts.footerHeight
}

export function contentWidthPx(opts: PaginationOptions): number {
  return pageDimension(opts).width - opts.margins.left - opts.margins.right
}
