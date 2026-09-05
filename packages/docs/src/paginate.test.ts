import { describe, expect, it } from 'vitest'
import { paginate, pageCount, pageOfBlock } from './paginate'
import { contentHeightPx, contentWidthPx, DEFAULT_PAGINATION_OPTIONS, type BlockBox, type LineBox, type PaginationOptions } from './types'

const letterOpts: PaginationOptions = DEFAULT_PAGINATION_OPTIONS
const CONTENT_H = contentHeightPx(letterOpts) // 11in - 2in margins = 9in = 864px

function block(id: string, height: number, breakBefore?: boolean): BlockBox {
  return { id, height, breakBefore }
}

/** count evenly-sized lines of lineHeight stacked from offset 0 — a block whose height is exactly count*lineHeight. */
function evenLines(count: number, lineHeight: number): LineBox[] {
  return Array.from({ length: count }, (_, i) => ({ offset: i * lineHeight, height: lineHeight }))
}
function lineBlock(id: string, count: number, lineHeight: number, breakBefore?: boolean): BlockBox {
  return { id, height: count * lineHeight, lineBoxes: evenLines(count, lineHeight), breakBefore }
}

describe('paginate', () => {
  it('empty document yields a single empty page', () => {
    const layout = paginate([], letterOpts)
    expect(pageCount(layout)).toBe(1)
    expect(layout.pages[0]).toEqual({ start: 0, end: 0 })
  })

  it('blocks that fit within one page stay on one page', () => {
    const blocks = [block('a', 100), block('b', 200), block('c', 300)]
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(1)
    expect(layout.pages[0]).toEqual({ start: 0, end: 3 })
  })

  it('overflow starts a new page at the block that would not fit', () => {
    // three blocks each just under half the content height: 2 fit, 3rd overflows
    const h = CONTENT_H / 2 - 10
    const blocks = [block('a', h), block('b', h), block('c', h)]
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[0]).toEqual({ start: 0, end: 2 })
    expect(layout.pages[1]).toEqual({ start: 2, end: 3 })
  })

  it('a block taller than a whole page still gets placed (page overflows visually)', () => {
    const blocks = [block('huge', CONTENT_H * 3)]
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(1)
    expect(layout.pages[0]).toEqual({ start: 0, end: 1 })
  })

  it('an oversized block does not swallow the block after it onto the same page', () => {
    const blocks = [block('huge', CONTENT_H * 1.5), block('next', 50)]
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[0]).toEqual({ start: 0, end: 1 })
    expect(layout.pages[1]).toEqual({ start: 1, end: 2 })
  })

  it('breakBefore forces a new page even when the block would otherwise fit', () => {
    const blocks = [block('a', 50), block('b', 50, true), block('c', 50)]
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[0]).toEqual({ start: 0, end: 1 })
    expect(layout.pages[1]).toEqual({ start: 1, end: 3 })
  })

  it('breakBefore on the very first block does not create a blank leading page', () => {
    const blocks = [block('a', 50, true), block('b', 50)]
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(1)
    expect(layout.pages[0]).toEqual({ start: 0, end: 2 })
  })

  it('multiple manual breaks in a row each start a fresh (possibly short) page', () => {
    const blocks = [block('a', 10), block('b', 10, true), block('c', 10, true)]
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(3)
    expect(layout.pages).toEqual([
      { start: 0, end: 1 },
      { start: 1, end: 2 },
      { start: 2, end: 3 },
    ])
  })

  it('respects a custom header/footer reservation shrinking content height', () => {
    const opts: PaginationOptions = { ...letterOpts, headerHeight: 40, footerHeight: 40 }
    const reduced = contentHeightPx(opts)
    expect(reduced).toBe(CONTENT_H - 80)
    const blocks = [block('a', reduced - 10), block('b', 20)]
    const layout = paginate(blocks, opts)
    expect(pageCount(layout)).toBe(2)
  })

  it('a4 pages are narrower but taller than letter at equal margins (real paper dimensions)', () => {
    const a4Opts: PaginationOptions = { ...letterOpts, pageSize: 'a4' }
    expect(contentWidthPx(a4Opts)).toBeLessThan(contentWidthPx(letterOpts))
    expect(contentHeightPx(a4Opts)).toBeGreaterThan(contentHeightPx(letterOpts))
  })

  it('pageOfBlock maps a block index to its page', () => {
    const h = CONTENT_H / 2 - 10
    const blocks = [block('a', h), block('b', h), block('c', h)]
    const layout = paginate(blocks, letterOpts)
    expect(pageOfBlock(layout, 0)).toBe(0)
    expect(pageOfBlock(layout, 1)).toBe(0)
    expect(pageOfBlock(layout, 2)).toBe(1)
  })
})

// ---- v2 (round 4): line-level splitting for real paper-container layout ----
describe('paginate: line-level splitting (v2)', () => {
  it('a block with lineBoxes that fits entirely is NOT split (no startOffset/endOffset)', () => {
    const blocks = [lineBlock('para', 5, 20)] // 100px, fits easily
    const layout = paginate(blocks, letterOpts)
    expect(pageCount(layout)).toBe(1)
    expect(layout.pages[0]).toEqual({ start: 0, end: 1 })
  })

  it('a block taller than the remaining space splits at the last line that fits', () => {
    // Fill the page to within 45px of the top, then a 10-line, 20px-per-line
    // (200px) paragraph: only 2 lines (40px) fit in the remaining 45px.
    const filler = block('filler', CONTENT_H - 45)
    const para = lineBlock('para', 10, 20)
    const layout = paginate([filler, para], letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[0]).toEqual({ start: 0, end: 2, endOffset: 40 })
    expect(layout.pages[1]).toEqual({ start: 1, end: 2, startOffset: 40 })
  })

  it('a block spanning THREE pages splits recursively, each continuation resuming exactly where the last left off', () => {
    const lineHeight = 20
    const linesPerPage = Math.floor(CONTENT_H / lineHeight)
    const totalLines = linesPerPage * 2 + 3 // spills a third of the way into a third page
    const para = lineBlock('giant', totalLines, lineHeight)
    const layout = paginate([para], letterOpts)
    expect(pageCount(layout)).toBe(3)

    const page1EndOffset = linesPerPage * lineHeight
    expect(layout.pages[0]).toEqual({ start: 0, end: 1, endOffset: page1EndOffset })
    expect(layout.pages[1]).toEqual({ start: 0, end: 1, startOffset: page1EndOffset, endOffset: page1EndOffset * 2 })
    expect(layout.pages[2]).toEqual({ start: 0, end: 1, startOffset: page1EndOffset * 2 })

    // No content lost or duplicated: the three offsets exactly cover [0, block height).
    expect(layout.pages[2].startOffset).toBeLessThan(para.height)
  })

  it('a block with lineBoxes but where not even one line fits falls back to whole-block-to-next-page (matches v1 oversized behavior)', () => {
    // filler leaves only 5px remaining; each line is 20px, so zero lines fit.
    const filler = block('filler', CONTENT_H - 5)
    const para = lineBlock('para', 4, 20)
    const layout = paginate([filler, para], letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[0]).toEqual({ start: 0, end: 1 })
    expect(layout.pages[1]).toEqual({ start: 1, end: 2 }) // whole block, no split fields
  })

  it('a lineBoxes block that is the FIRST thing on its page and still does not fit any line gets placed anyway (page overflows visually)', () => {
    // A single, absurdly tall single-line "paragraph" (e.g. one giant unbreakable word) as page 1's only content.
    const para: BlockBox = { id: 'huge-line', height: CONTENT_H * 2, lineBoxes: [{ offset: 0, height: CONTENT_H * 2 }] }
    const layout = paginate([para], letterOpts)
    expect(pageCount(layout)).toBe(1)
    expect(layout.pages[0]).toEqual({ start: 0, end: 1 })
  })

  it('breakBefore still forces a fresh page even for a lineBoxes-carrying block that would otherwise fit', () => {
    const a = lineBlock('a', 3, 20)
    const b = lineBlock('b', 3, 20, true)
    const layout = paginate([a, b], letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[0]).toEqual({ start: 0, end: 1 })
    expect(layout.pages[1]).toEqual({ start: 1, end: 2 })
  })

  it('a split block followed by a normal block: the normal block correctly starts fresh (no stray offset fields)', () => {
    const filler = block('filler', CONTENT_H - 45)
    const para = lineBlock('para', 10, 20) // splits: 2 lines page1, 8 lines page2
    const after = block('after', 50)
    const layout = paginate([filler, para, after], letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[1]).toEqual({ start: 1, end: 3, startOffset: 40 })
  })

  it('mixing plain blocks (no lineBoxes) and lineBoxes blocks in the same document works correctly', () => {
    const heading = block('h1', 40) // no lineBoxes — v1 whole-block
    const filler = block('filler', CONTENT_H - 40 - 45)
    const para = lineBlock('para', 10, 20)
    const layout = paginate([heading, filler, para], letterOpts)
    expect(pageCount(layout)).toBe(2)
    expect(layout.pages[0].end).toBe(3)
    expect(layout.pages[0].endOffset).toBe(40)
  })
})
