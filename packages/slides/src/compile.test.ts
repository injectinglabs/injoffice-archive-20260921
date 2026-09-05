import { describe, expect, it } from 'vitest'
import { compileDeckToWire, compileSlide, splitInlineRuns } from './compile'
import { boardroomTheme } from './themes'
import type { DeckSpec, SlideSpec } from './types'

function slide(partial: Partial<SlideSpec> & { kind: SlideSpec['kind'] }, id = 's1'): SlideSpec {
  return { id, ...partial }
}

describe('splitInlineRuns', () => {
  it('returns nothing for empty text', () => {
    expect(splitInlineRuns('', {})).toEqual([])
  })

  it('is a single plain run for plain text', () => {
    const runs = splitInlineRuns('hello world', { sizePt: 10 })
    expect(runs).toEqual([{ text: 'hello world', sizePt: 10 }])
  })

  it('splits bold and italic markers into styled runs', () => {
    const runs = splitInlineRuns('a **bold** and *italic* word', {})
    expect(runs.map((r) => r.text)).toEqual(['a ', 'bold', ' and ', 'italic', ' word'])
    expect(runs[1]?.bold).toBe(true)
    expect(runs[3]?.italic).toBe(true)
    expect(runs[0]?.bold).toBeUndefined()
  })

  it('preserves base style on every split run', () => {
    const runs = splitInlineRuns('plain **bold**', { color: '#112233', font: 'Archivo' })
    for (const r of runs) {
      expect(r.color).toBe('#112233')
      expect(r.font).toBe('Archivo')
    }
  })

  it('base bold/italic stays true even without a marker', () => {
    const runs = splitInlineRuns('already bold', { bold: true })
    expect(runs[0]?.bold).toBe(true)
  })
})

describe('compileSlide', () => {
  it('title slide: title placeholder + accent bar, no page number', () => {
    const s = slide({ kind: 'title', title: 'Hello', subtitle: 'World' })
    const wire = compileSlide(s, boardroomTheme, 0, 3)
    const titleShape = wire.shapes.find((sh) => sh.placeholder === 'title')
    expect(titleShape).toBeDefined()
    expect(titleShape?.paragraphs?.[0]?.runs[0]?.text).toBe('Hello')
    expect(wire.background).toBe(boardroomTheme.background)
    // title kind never gets the page-number chip, even with total > 1
    expect(wire.shapes.every((sh) => !sh.paragraphs?.some((p) => p.runs.some((r) => r.text.includes('/'))))).toBe(true)
  })

  it('bullets slide: body placeholder carries real bulleted paragraphs', () => {
    const s = slide({ kind: 'bullets', title: 'Points', bullets: ['One', 'Two', 'Three'] })
    const wire = compileSlide(s, boardroomTheme, 1, 3)
    const body = wire.shapes.find((sh) => sh.placeholder === 'body')
    expect(body).toBeDefined()
    expect(body?.paragraphs?.length).toBe(3)
    for (const p of body?.paragraphs ?? []) {
      expect(p.bullet).toBe(true)
    }
    expect(body?.paragraphs?.[1]?.runs[0]?.text).toBe('Two')
  })

  it('non-title slide with more than one total slide gets a page-number chip', () => {
    const s = slide({ kind: 'bullets', title: 'Points', bullets: ['One'] })
    const wire = compileSlide(s, boardroomTheme, 1, 4)
    const found = wire.shapes.some((sh) => sh.paragraphs?.some((p) => p.runs.some((r) => r.text === '2 / 4')))
    expect(found).toBe(true)
  })

  it('a single-slide deck gets no page-number chip', () => {
    const s = slide({ kind: 'bullets', title: 'Points', bullets: ['One'] })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    const found = wire.shapes.some((sh) => sh.paragraphs?.some((p) => p.runs.some((r) => r.text.includes('/'))))
    expect(found).toBe(false)
  })

  it('two-col slide compiles both columns with real bullet paragraphs and column titles', () => {
    const s = slide({
      kind: 'two-col',
      title: 'Compare',
      colTitles: ['Before', 'After'],
      bullets: ['Slow'],
      bulletsRight: ['Fast'],
    })
    const wire = compileSlide(s, boardroomTheme, 0, 2)
    const texts = wire.shapes.flatMap((sh) => sh.paragraphs?.flatMap((p) => p.runs.map((r) => r.text)) ?? [])
    expect(texts).toContain('BEFORE')
    expect(texts).toContain('AFTER')
    expect(texts).toContain('Slow')
    expect(texts).toContain('Fast')
  })

  it('quote slide wraps the quote in curly quotes and uppercases the attribution', () => {
    const s = slide({ kind: 'quote', quote: 'Stay hungry', subtitle: 'Steve' })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    const texts = wire.shapes.flatMap((sh) => sh.paragraphs?.flatMap((p) => p.runs.map((r) => r.text)) ?? [])
    expect(texts.some((t) => t.includes('“Stay hungry”'))).toBe(true)
    expect(texts.some((t) => t === '— STEVE')).toBe(true)
  })

  it('closing slide defaults its title to "Thank you" and centers text', () => {
    const s = slide({ kind: 'closing' })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    const titleShape = wire.shapes.find((sh) => sh.placeholder === 'ctrTitle')
    expect(titleShape?.paragraphs?.[0]?.runs[0]?.text).toBe('Thank you')
    expect(titleShape?.paragraphs?.[0]?.align).toBe('ctr')
  })

  it('geometry is in EMU and matches the 13.333x7.5in 16:9 canvas', () => {
    const s = slide({ kind: 'title', title: 'X' })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    const title = wire.shapes.find((sh) => sh.placeholder === 'title')!
    // 0.9in = 823,... EMU at 914400/in — exact integer.
    expect(title.x).toBe(Math.round(0.9 * 914400))
    expect(title.cx).toBe(Math.round(10.8 * 914400))
  })
})

describe('shape keys (S8 persistence identity)', () => {
  it('every shape on every kind gets a stable, non-empty key', () => {
    const kinds: SlideSpec['kind'][] = ['title', 'section', 'bullets', 'two-col', 'quote', 'closing']
    for (const kind of kinds) {
      const s = slide({ kind, title: 'T', subtitle: 'S', body: 'B', quote: 'Q', bullets: ['a'], bulletsRight: ['b'], eyebrow: 'E' })
      const wire = compileSlide(s, boardroomTheme, 1, 3) // index>0, total>1 so page-number compiles too
      for (const shape of wire.shapes) {
        expect(shape.key, `${kind} shape ${shape.name ?? shape.kind} should have a key`).toBeTruthy()
      }
    }
  })

  it('keys are stable across content edits that do not add/remove shapes', () => {
    const a = compileSlide(slide({ kind: 'bullets', title: 'One', bullets: ['x'] }), boardroomTheme, 0, 1)
    const b = compileSlide(slide({ kind: 'bullets', title: 'Two', bullets: ['x', 'y'] }), boardroomTheme, 0, 1)
    const keysOf = (wire: typeof a) => wire.shapes.map((s) => s.key).sort()
    expect(keysOf(a)).toEqual(keysOf(b))
  })
})

describe('shapeOverrides (S8 persistence)', () => {
  it('overrides the position/size of the shape with a matching key, leaves others untouched', () => {
    const s = slide({ kind: 'title', title: 'Hello', subtitle: 'World' })
    const withOverride: SlideSpec = { ...s, shapeOverrides: { title: { x: 111, y: 222, cx: 333, cy: 444 } } }
    const wire = compileSlide(withOverride, boardroomTheme, 0, 1)
    const title = wire.shapes.find((sh) => sh.key === 'title')!
    expect(title).toEqual(expect.objectContaining({ x: 111, y: 222, cx: 333, cy: 444 }))
    const subtitle = wire.shapes.find((sh) => sh.key === 'subtitle')!
    const baseline = compileSlide(s, boardroomTheme, 0, 1).shapes.find((sh) => sh.key === 'subtitle')!
    expect(subtitle).toEqual(baseline)
  })

  it('a partial override only replaces the fields it sets', () => {
    const s = slide({ kind: 'title', title: 'Hello' })
    const baseline = compileSlide(s, boardroomTheme, 0, 1).shapes.find((sh) => sh.key === 'title')!
    const withOverride: SlideSpec = { ...s, shapeOverrides: { title: { x: 999 } } }
    const wire = compileSlide(withOverride, boardroomTheme, 0, 1)
    const title = wire.shapes.find((sh) => sh.key === 'title')!
    expect(title.x).toBe(999)
    expect(title.y).toBe(baseline.y)
    expect(title.cx).toBe(baseline.cx)
    expect(title.cy).toBe(baseline.cy)
  })

  it('an override whose key matches nothing compiled is silently inert', () => {
    const s = slide({ kind: 'title', title: 'Hello' }) // no subtitle, no body
    const withOverride: SlideSpec = { ...s, shapeOverrides: { subtitle: { x: 1, y: 2, cx: 3, cy: 4 }, madeUp: { x: 5 } } }
    const wire = compileSlide(withOverride, boardroomTheme, 0, 1)
    const baseline = compileSlide(s, boardroomTheme, 0, 1)
    expect(wire).toEqual(baseline)
  })

  it('a deck with no shapeOverrides at all compiles identically to before this feature existed', () => {
    const s = slide({ kind: 'bullets', title: 'Points', bullets: ['a', 'b'] })
    const withEmptyField: SlideSpec = { ...s, shapeOverrides: undefined }
    expect(compileSlide(withEmptyField, boardroomTheme, 0, 1)).toEqual(compileSlide(s, boardroomTheme, 0, 1))
  })

  it('a slide with no diagram field compiles identically to before S12 existed', () => {
    const s = slide({ kind: 'bullets', title: 'Points', bullets: ['a'] })
    const withUndefinedDiagram: SlideSpec = { ...s, diagram: undefined }
    expect(compileSlide(withUndefinedDiagram, boardroomTheme, 0, 1)).toEqual(compileSlide(s, boardroomTheme, 0, 1))
  })

  it('a slide with a diagram appends the diagram shapes on top of its own kind-produced shapes', () => {
    const s = slide({ kind: 'bullets', title: 'Pipeline', diagram: { type: 'process', steps: ['A', 'B', 'C'] } })
    const withoutDiagram = slide({ kind: 'bullets', title: 'Pipeline' })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    const baseline = compileSlide(withoutDiagram, boardroomTheme, 0, 1)
    // Every baseline shape (title, empty bullets body) is still present...
    expect(wire.shapes.length).toBeGreaterThan(baseline.shapes.length)
    // ...plus 3 process-flow boxes and 2 connectors.
    expect(wire.shapes.filter((sh) => sh.key?.startsWith('diagram-step-')).length).toBe(3)
    expect(wire.shapes.filter((sh) => sh.key?.startsWith('diagram-arrow-')).length).toBe(2)
  })

  it('an org-chart diagram on a title-less slide starts higher (no title band to clear)', () => {
    const s = slide({ kind: 'bullets', diagram: { type: 'orgChart', root: { label: 'CEO', children: [{ label: 'CTO' }] } } })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    const root = wire.shapes.find((sh) => sh.key === 'diagram-org-0')
    expect(root).toBeDefined()
    expect(root!.y).toBeLessThan(inchToEmu(1.5)) // well above the y=2.3in a titled slide would use
  })
})

function inchToEmu(n: number): number {
  return Math.round(n * 914400)
}

describe('shapeAnimations + transition (S11)', () => {
  it('applies a fade entrance to the shape with a matching key, leaves others untouched', () => {
    const s = slide({ kind: 'title', title: 'Hello', subtitle: 'World' })
    const withAnim: SlideSpec = { ...s, shapeAnimations: { title: { enter: 'fade', delayMs: 200 } } }
    const wire = compileSlide(withAnim, boardroomTheme, 0, 1)
    const title = wire.shapes.find((sh) => sh.key === 'title')!
    expect(title.enter).toBe('fade')
    expect(title.enterDelayMs).toBe(200)
    const subtitle = wire.shapes.find((sh) => sh.key === 'subtitle')!
    expect(subtitle.enter).toBeUndefined()
  })

  it('applies a flyIn entrance with direction/distance', () => {
    const s = slide({ kind: 'title', title: 'Hello' })
    const withAnim: SlideSpec = { ...s, shapeAnimations: { title: { enter: 'flyIn', direction: 'left', distance: 0.4 } } }
    const wire = compileSlide(withAnim, boardroomTheme, 0, 1)
    const title = wire.shapes.find((sh) => sh.key === 'title')!
    expect(title.enter).toBe('flyIn')
    expect(title.enterDirection).toBe('left')
    expect(title.enterDistance).toBe(0.4)
  })

  it('an animation whose key matches nothing compiled is silently inert', () => {
    const s = slide({ kind: 'title', title: 'Hello' }) // no subtitle, no body
    const withAnim: SlideSpec = { ...s, shapeAnimations: { subtitle: { enter: 'fade' }, madeUp: { enter: 'flyIn' } } }
    const wire = compileSlide(withAnim, boardroomTheme, 0, 1)
    const baseline = compileSlide(s, boardroomTheme, 0, 1)
    expect(wire).toEqual(baseline)
  })

  it('a deck with no shapeAnimations/transition at all compiles identically to before this feature existed', () => {
    const s = slide({ kind: 'bullets', title: 'Points', bullets: ['a', 'b'] })
    const withEmptyFields: SlideSpec = { ...s, shapeAnimations: undefined, transition: undefined }
    expect(compileSlide(withEmptyFields, boardroomTheme, 0, 1)).toEqual(compileSlide(s, boardroomTheme, 0, 1))
  })

  it('sets a slide-level transition kind and direction on the wire slide', () => {
    const s = slide({ kind: 'title', title: 'Hello', transition: { kind: 'push', direction: 'left' } })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    expect(wire.transition).toBe('push')
    expect(wire.transitionDirection).toBe('left')
  })

  it('omits transitionDirection when the transition spec has none (fade)', () => {
    const s = slide({ kind: 'title', title: 'Hello', transition: { kind: 'fade' } })
    const wire = compileSlide(s, boardroomTheme, 0, 1)
    expect(wire.transition).toBe('fade')
    expect(wire.transitionDirection).toBeUndefined()
  })
})

describe('compileDeckToWire', () => {
  it('compiles every slide and carries the resolved theme background through', () => {
    const deck: DeckSpec = {
      id: 'd1',
      title: 'Deck',
      theme: 'midnight',
      slides: [slide({ kind: 'title', id: 'a', title: 'Hi' }), slide({ kind: 'bullets', id: 'b', title: 'Pts', bullets: ['x'] })],
    }
    const wire = compileDeckToWire(deck)
    expect(wire.slides).toHaveLength(2)
    expect(wire.slides[0]?.background).toBe('#131a1b') // midnight theme background
    expect(wire.slides[1]?.background).toBe('#131a1b')
  })

  it('an empty deck compiles to an empty slide list without throwing', () => {
    const deck: DeckSpec = { id: 'd2', title: 'Empty', slides: [] }
    expect(compileDeckToWire(deck)).toEqual({ slides: [] })
  })
})
