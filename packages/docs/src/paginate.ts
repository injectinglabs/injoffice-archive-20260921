import { contentHeightPx, type BlockBox, type PageLayout, type PageSlice, type PaginationOptions } from './types'

const EPSILON = 1e-7

interface OpenPage {
  start: number
  end: number
  startOffset?: number
  endOffset?: number
  used: number
}

function checkedExtent(value: number, label: string): number {
  if (!Number.isFinite(value) || value < 0) throw new RangeError(`${label} must be a finite non-negative number`)
  return value
}

function lineEnds(block: BlockBox): number[] {
  if (!block.lineBoxes?.length) return []
  const ends: number[] = []
  let previousTop = -Infinity
  let previousEnd = -Infinity

  for (const line of block.lineBoxes) {
    const top = checkedExtent(line.offset, `line offset for block ${block.id}`)
    const height = checkedExtent(line.height, `line height for block ${block.id}`)
    const end = top + height
    if (top + EPSILON < previousTop || end + EPSILON < previousEnd) {
      throw new RangeError(`line boxes for block ${block.id} must be in ascending order`)
    }
    if (end > previousEnd + EPSILON) ends.push(end)
    previousTop = top
    previousEnd = Math.max(previousEnd, end)
  }

  return ends
}

function closePage(open: OpenPage): PageSlice {
  const page: PageSlice = { start: open.start, end: open.end }
  if (open.startOffset !== undefined && open.startOffset > EPSILON) page.startOffset = open.startOffset
  if (open.endOffset !== undefined) page.endOffset = open.endOffset
  return page
}

function beginsEmpty(index: number): OpenPage {
  return { start: index, end: index, used: 0 }
}

/** Lay measured blocks onto fixed-height page content boxes. */
export function paginate(blocks: BlockBox[], opts: PaginationOptions): PageLayout {
  const capacity = checkedExtent(contentHeightPx(opts), 'page content height')
  const pages: PageSlice[] = []
  let open = beginsEmpty(0)

  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex]!
    const blockHeight = checkedExtent(block.height, `height for block ${block.id}`)
    const boundaries = lineEnds(block).filter((end) => end <= blockHeight + EPSILON)
    let offset = 0

    if (block.breakBefore && open.end > open.start) {
      pages.push(closePage(open))
      open = beginsEmpty(blockIndex)
    }

    while (offset < blockHeight - EPSILON || (blockHeight === 0 && offset === 0)) {
      const fragmentHeight = blockHeight - offset
      const remaining = capacity - open.used

      if (fragmentHeight <= remaining + EPSILON || fragmentHeight === 0) {
        if (open.end === open.start && offset > EPSILON) open.startOffset = offset
        open.end = blockIndex + 1
        open.endOffset = undefined
        open.used += Math.max(0, fragmentHeight)
        offset = blockHeight
        break
      }

      const fitting = boundaries.filter((end) => end > offset + EPSILON && end - offset <= remaining + EPSILON)
      if (fitting.length > 0) {
        const splitAt = fitting[fitting.length - 1]!
        if (open.end === open.start && offset > EPSILON) open.startOffset = offset
        open.end = blockIndex + 1
        open.endOffset = splitAt
        open.used += splitAt - offset
        pages.push(closePage(open))
        offset = splitAt
        open = beginsEmpty(blockIndex)
        open.startOffset = offset
        continue
      }

      if (open.end > open.start) {
        pages.push(closePage(open))
        open = beginsEmpty(blockIndex)
        if (offset > EPSILON) open.startOffset = offset
        continue
      }

      const nextBoundary = boundaries.find((end) => end > offset + EPSILON)
      if (nextBoundary !== undefined && nextBoundary < blockHeight - EPSILON) {
        open.startOffset = offset > EPSILON ? offset : undefined
        open.end = blockIndex + 1
        open.endOffset = nextBoundary
        open.used = nextBoundary - offset
        pages.push(closePage(open))
        offset = nextBoundary
        open = beginsEmpty(blockIndex)
        open.startOffset = offset
        continue
      }

      // No legal split exists. Put the complete remaining fragment on this
      // otherwise-empty page, even when it overflows, so progress is certain.
      open.startOffset = offset > EPSILON ? offset : undefined
      open.end = blockIndex + 1
      open.endOffset = undefined
      open.used = fragmentHeight
      offset = blockHeight
    }
  }

  if (blocks.length === 0 || open.end > open.start) pages.push(closePage(open))
  return { pages, contentHeight: capacity }
}

/** How many pages the layout occupies. */
export function pageCount(layout: PageLayout): number {
  return layout.pages.length
}

/** Return the first zero-based page containing the requested block index. */
export function pageOfBlock(layout: PageLayout, blockIndex: number): number {
  if (layout.pages.length === 0) return 0
  for (let page = 0; page < layout.pages.length; page++) {
    const slice = layout.pages[page]!
    if (blockIndex >= slice.start && blockIndex < slice.end) return page
  }
  return blockIndex < layout.pages[0]!.start ? 0 : layout.pages.length - 1
}
