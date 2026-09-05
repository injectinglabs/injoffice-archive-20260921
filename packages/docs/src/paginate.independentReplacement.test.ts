import { describe, expect, it } from 'vitest'
import { pageCount, pageOfBlock, paginate } from './paginate'
import type { PaginationOptions } from './types'

const OPTIONS: PaginationOptions = {
  pageSize: 'letter',
  margins: { top: 478, bottom: 478, left: 96, right: 96 },
  headerHeight: 0,
  footerHeight: 0,
}

describe('independently authored measured-block pagination fixtures', () => {
  it('keeps one editable sheet for an empty story', () => {
    expect(paginate([], OPTIONS)).toEqual({ pages: [{ start: 0, end: 0 }], contentHeight: 100 })
  })

  it('moves an atomic block and guarantees progress for an oversized one', () => {
    const layout = paginate([
      { id: 'intro', height: 70 },
      { id: 'card', height: 50 },
      { id: 'poster', height: 180 },
    ], OPTIONS)
    expect(layout.pages).toEqual([
      { start: 0, end: 1 },
      { start: 1, end: 2 },
      { start: 2, end: 3 },
    ])
  })

  it('splits solely at measured line bottoms and preserves source indices', () => {
    const layout = paginate([
      { id: 'lead', height: 60 },
      {
        id: 'paragraph',
        height: 100,
        lineBoxes: [0, 25, 50, 75].map((offset) => ({ offset, height: 25 })),
      },
    ], OPTIONS)
    expect(layout.pages).toEqual([
      { start: 0, end: 2, endOffset: 25 },
      { start: 1, end: 2, startOffset: 25 },
    ])
    expect(pageCount(layout)).toBe(2)
    expect(pageOfBlock(layout, 1)).toBe(0)
    expect(pageOfBlock(layout, 20)).toBe(1)
  })

  it('honors break-before without manufacturing a leading blank page', () => {
    const layout = paginate([
      { id: 'first', height: 20, breakBefore: true },
      { id: 'second', height: 20, breakBefore: true },
      { id: 'third', height: 20, breakBefore: true },
    ], OPTIONS)
    expect(layout.pages).toEqual([
      { start: 0, end: 1 },
      { start: 1, end: 2 },
      { start: 2, end: 3 },
    ])
  })

  it('places one over-height line per page rather than stalling', () => {
    const layout = paginate([
      { id: 'large-type', height: 240, lineBoxes: [{ offset: 0, height: 120 }, { offset: 120, height: 120 }] },
    ], OPTIONS)
    expect(layout.pages).toEqual([
      { start: 0, end: 1, endOffset: 120 },
      { start: 0, end: 1, startOffset: 120 },
    ])
  })
})
