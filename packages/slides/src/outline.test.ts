import { beforeEach, describe, expect, it } from 'vitest'
import { deckFromOutline, resetSlideIds } from './outline'
import { resolveTheme, boardroomTheme, midnightTheme } from './themes'

beforeEach(() => resetSlideIds())

describe('deckFromOutline', () => {
  it('maps the outline grammar to slide kinds', () => {
    const deck = deckFromOutline('Q3 Review', [
      '# Q3 Review',
      'What happened and what we do next',
      '## Highlights',
      '- Revenue **up 40%**',
      '- Churn flat',
      'Notes: pause here for questions',
      '> We stopped guessing — the numbers decide. — Ops lead',
      '# Next quarter',
      '## Bets',
      '- Bet 1',
    ].join('\n'))
    const kinds = deck.slides.map((s) => s.kind)
    expect(kinds).toEqual(['title', 'bullets', 'quote', 'section', 'bullets'])
    expect(deck.slides[0]).toMatchObject({ title: 'Q3 Review', subtitle: 'What happened and what we do next' })
    expect(deck.slides[1].bullets).toEqual(['Revenue **up 40%**', 'Churn flat'])
    expect(deck.slides[1].notes).toBe('pause here for questions')
    expect(deck.slides[2]).toMatchObject({ quote: 'We stopped guessing — the numbers decide.', subtitle: 'Ops lead' })
    expect(deck.slides[3].kind).toBe('section')
  })

  it('never drops content: stray text and lone bullets get slides', () => {
    const deck = deckFromOutline('T', 'just a line\n- lone bullet')
    expect(deck.slides[0]).toMatchObject({ kind: 'title', title: 'T', subtitle: 'just a line' })
    expect(deck.slides[1]).toMatchObject({ kind: 'bullets', bullets: ['lone bullet'] })
  })

  it('empty outline yields a single title slide', () => {
    const deck = deckFromOutline('Solo', '')
    expect(deck.slides).toHaveLength(1)
    expect(deck.slides[0]).toMatchObject({ kind: 'title', title: 'Solo' })
  })
})

describe('resolveTheme', () => {
  it('knows every built-in theme id', () => {
    // Keep in sync with the gateway's python generator THEMES dict.
    const ids = ['boardroom', 'midnight', 'slate', 'terra', 'forest', 'plum']
    for (const id of ids) expect(resolveTheme(id).id).toBe(id)
  })
  it('resolves ids, objects and defaults', () => {
    expect(resolveTheme(undefined)).toBe(boardroomTheme)
    expect(resolveTheme('midnight')).toBe(midnightTheme)
    expect(resolveTheme('nope')).toBe(boardroomTheme)
    const custom = resolveTheme({ ...boardroomTheme, id: 'brand', accent: '#ff0055' })
    expect(custom.accent).toBe('#ff0055')
    expect(custom.displayFont).toBe(boardroomTheme.displayFont)
  })
})
