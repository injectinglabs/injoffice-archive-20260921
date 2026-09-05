import { describe, expect, it } from 'vitest'
import { addSlide, bulletsToText, moveSlide, removeSlide, textToBullets, updateSlide } from './edit'
import type { DeckSpec } from './types'

const deck = (): DeckSpec => ({
  id: 'd',
  title: 'T',
  slides: [
    { id: 'a', kind: 'title', title: 'A' },
    { id: 'b', kind: 'bullets', title: 'B', bullets: ['one'] },
    { id: 'c', kind: 'closing', title: 'C' },
  ],
})

describe('deck edit operations', () => {
  it('updateSlide patches immutably and preserves the id', () => {
    const d = deck()
    const next = updateSlide(d, 1, { title: 'B2', id: 'hax' } as never)
    expect(next.slides[1]).toMatchObject({ id: 'b', title: 'B2', bullets: ['one'] })
    expect(d.slides[1].title).toBe('B')
    expect(updateSlide(d, 9, { title: 'x' })).toBe(d)
  })
  it('addSlide inserts after the index with a kind-appropriate blank and fresh id', () => {
    const d = deck()
    const next = addSlide(d, 0, 'quote')
    expect(next.slides).toHaveLength(4)
    expect(next.slides[1].kind).toBe('quote')
    expect(next.slides[1].quote).toBeTruthy()
    expect(new Set(next.slides.map((s) => s.id)).size).toBe(4)
    expect(addSlide(d, -1, 'bullets').slides[3].kind).toBe('bullets')
  })
  it('removeSlide keeps at least one slide', () => {
    const d = deck()
    expect(removeSlide(d, 1).slides.map((s) => s.id)).toEqual(['a', 'c'])
    const single: DeckSpec = { id: 'd', title: 'T', slides: [{ id: 'a', kind: 'title' }] }
    expect(removeSlide(single, 0)).toBe(single)
  })
  it('moveSlide clamps and reorders', () => {
    const d = deck()
    expect(moveSlide(d, 0, 1).slides.map((s) => s.id)).toEqual(['b', 'a', 'c'])
    expect(moveSlide(d, 0, -1)).toBe(d)
    expect(moveSlide(d, 2, 1)).toBe(d)
  })
  it('bullets round-trip through the text block form', () => {
    expect(textToBullets('one\n- two\n*  three\n\n')).toEqual(['one', 'two', 'three'])
    expect(textToBullets('  \n')).toBeUndefined()
    expect(bulletsToText(['a', 'b'])).toBe('a\nb')
    expect(bulletsToText(undefined)).toBe('')
  })
})
