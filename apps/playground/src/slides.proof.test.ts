import { describe, expect, it } from 'vitest'
import { deckFromOutline, resetSlideIds } from '../../../packages/slides/src/outline'
import { PRESENTATION_DEMO_OUTLINE, PRESENTATION_DEMO_TITLE } from './presentationDemoFixtures'

describe('slides DeckSpec proof', () => {
  it('builds at least two slides from an outline', () => {
    resetSlideIds()
    const deck = deckFromOutline('InjOffice slides', '# Title\n## Body\n- one\n')
    expect(deck.slides.length).toBeGreaterThanOrEqual(2)
  })

  it('ships a useful multi-slide starting outline instead of an empty deck', () => {
    resetSlideIds()
    const deck = deckFromOutline(PRESENTATION_DEMO_TITLE, PRESENTATION_DEMO_OUTLINE)
    expect(deck.slides).toHaveLength(6)
    expect(deck.slides.some((slide) => slide.kind === 'quote')).toBe(true)
    expect(deck.slides.flatMap((slide) => slide.bullets ?? []).length).toBeGreaterThanOrEqual(10)
  })
})
