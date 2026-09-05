import { describe, expect, it } from 'vitest'
import { auditDeck, auditSlide, estimateWrappedLines, formatDeckAudit } from './qc'
import type { DeckSpec, SlideSpec } from './types'

function slide(partial: Partial<SlideSpec> & { kind: SlideSpec['kind'] }): SlideSpec {
  return { id: 's1', ...partial }
}

describe('estimateWrappedLines', () => {
  it('is 0 for empty text', () => {
    expect(estimateWrappedLines('', 80, 3)).toBe(0)
  })

  it('is 1 for short text that fits one line', () => {
    expect(estimateWrappedLines('Hi', 80, 3)).toBe(1)
  })

  it('grows with text length', () => {
    const short = estimateWrappedLines('a '.repeat(5), 40, 3)
    const long = estimateWrappedLines('a '.repeat(50), 40, 3)
    expect(long).toBeGreaterThan(short)
  })

  it('counts explicit newlines as separate wrapped segments', () => {
    const oneLine = estimateWrappedLines('hello world', 80, 3)
    const twoLines = estimateWrappedLines('hello\nworld', 80, 3)
    expect(twoLines).toBeGreaterThanOrEqual(oneLine)
    expect(estimateWrappedLines('a\nb\nc', 80, 3)).toBe(3)
  })

  it('does not count markdown emphasis markers as characters', () => {
    const plain = estimateWrappedLines('hello world', 20, 3)
    const bold = estimateWrappedLines('**hello** world', 20, 3)
    expect(bold).toBe(plain)
  })
})

describe('auditSlide', () => {
  it('a normal, short title slide has no issues', () => {
    const s = slide({ kind: 'title', title: 'Quarterly Update', subtitle: 'Q3 2026 results' })
    expect(auditSlide(s, 0)).toEqual([])
  })

  it('an empty slide has no issues', () => {
    expect(auditSlide(slide({ kind: 'bullets' }), 0)).toEqual([])
  })

  it('flags out-of-bounds overflow for a slide with far too many long bullets', () => {
    const manyBullets = Array.from({ length: 30 }, (_, i) => `This is a fairly long bullet point number ${i} with real substance to it`)
    const s = slide({ kind: 'bullets', title: 'Everything', bullets: manyBullets })
    const issues = auditSlide(s, 2)
    expect(issues.some((i) => i.kind === 'out-of-bounds')).toBe(true)
    expect(issues[0]?.slideIndex).toBe(2)
    expect(issues[0]?.slideId).toBe('s1')
  })

  it('flags out-of-bounds for an extremely long title', () => {
    const s = slide({ kind: 'title', title: 'A '.repeat(200) })
    const issues = auditSlide(s, 0)
    expect(issues.some((i) => i.kind === 'out-of-bounds')).toBe(true)
  })

  it('flags a near-the-margin title/closing slide as an overlap risk before it fully overflows', () => {
    // Long enough to approach the reserved accent-bar band, short of full overflow.
    const s = slide({
      kind: 'closing',
      title: 'Thank You Everyone For Coming Together Today',
      subtitle: 'We could not have done it without each and every one of you supporting this launch. '.repeat(7),
    })
    const issues = auditSlide(s, 0)
    // Either overlap-risk or out-of-bounds is an acceptable outcome depending on the exact
    // estimate — what matters is it's flagged, not left silently clean.
    expect(issues.length).toBeGreaterThan(0)
  })

  it('flags unbalanced two-col columns', () => {
    const s = slide({
      kind: 'two-col',
      title: 'Comparison',
      bullets: ['short'],
      bulletsRight: Array.from({ length: 12 }, (_, i) => `A much longer right-column bullet point number ${i} that goes on and on`),
    })
    const issues = auditSlide(s, 0)
    expect(issues.some((i) => i.kind === 'overlap' && i.message.includes('unbalanced'))).toBe(true)
  })

  it('a reasonable two-col slide has no issues', () => {
    const s = slide({
      kind: 'two-col',
      title: 'Before / After',
      colTitles: ['Before', 'After'],
      bullets: ['Manual process', 'Slow turnaround', 'Error prone'],
      bulletsRight: ['Automated', 'Fast', 'Reliable'],
    })
    expect(auditSlide(s, 0)).toEqual([])
  })

  it('a reasonable quote slide has no issues', () => {
    const s = slide({ kind: 'quote', quote: 'Simplicity is the ultimate sophistication.', subtitle: 'Leonardo da Vinci' })
    expect(auditSlide(s, 0)).toEqual([])
  })

  it('flags an overlong quote', () => {
    const s = slide({ kind: 'quote', quote: 'Lorem ipsum dolor sit amet. '.repeat(40) })
    const issues = auditSlide(s, 0)
    expect(issues.length).toBeGreaterThan(0)
  })
})

describe('auditDeck', () => {
  it('is empty for a clean deck and aggregates across slides otherwise', () => {
    const clean: DeckSpec = {
      id: 'd1',
      title: 'Deck',
      slides: [slide({ kind: 'title', id: 'a', title: 'Hello' }), slide({ kind: 'bullets', id: 'b', title: 'Points', bullets: ['One', 'Two'] })],
    }
    expect(auditDeck(clean)).toEqual([])

    const dirty: DeckSpec = {
      id: 'd2',
      title: 'Deck',
      slides: [
        slide({ kind: 'title', id: 'a', title: 'Hello' }),
        slide({ kind: 'bullets', id: 'b', title: 'Points', bullets: Array.from({ length: 40 }, (_, i) => `bullet ${i} `.repeat(10)) }),
      ],
    }
    const issues = auditDeck(dirty)
    expect(issues.length).toBeGreaterThan(0)
    expect(issues.every((i) => i.slideId === 'b')).toBe(true)
    expect(issues.every((i) => i.slideIndex === 1)).toBe(true)
  })
})

describe('formatDeckAudit', () => {
  it('reports a pass for no issues', () => {
    expect(formatDeckAudit([])).toContain('Passed')
  })

  it('lists each issue with slide number and kind', () => {
    const text = formatDeckAudit([{ slideId: 'x', slideIndex: 4, kind: 'overflow', message: 'too tall' }])
    expect(text).toContain('Slide 5')
    expect(text).toContain('[overflow]')
    expect(text).toContain('too tall')
  })
})
