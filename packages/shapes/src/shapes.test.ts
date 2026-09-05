import { describe, expect, it } from 'vitest'
import { hasText, SHAPE_CATEGORIES, SHAPE_DEFAULTS, SHAPE_KINDS } from './types'
import { shapePreviewFidelity } from './ShapeFloat'

// Geometry path-math tests live in geometry.test.ts. This file covers the
// type-catalogue's own internal consistency — the kind of bug that's easy
// to introduce with an 82-entry vocabulary (a category listing a kind that
// SHAPE_DEFAULTS forgot, a typo that silently drops a preset).

describe('shape catalogue', () => {
  it('every kind has a defaults entry and a well-typed text answer', () => {
    expect(SHAPE_KINDS.length).toBe(121) // 119 curated presets + line + text
    for (const k of SHAPE_KINDS) {
      expect(SHAPE_DEFAULTS[k]).toBeDefined()
      expect(typeof hasText(k)).toBe('boolean')
    }
  })

  it('line has no text; a basic shape and a callout both do', () => {
    expect(hasText('line')).toBe(false)
    expect(hasText('rect')).toBe(true)
    expect(hasText('wedgeRectCallout')).toBe(true)
  })

  it('the arrow family carries no text (rendered as a solid glyph)', () => {
    expect(hasText('rightArrow')).toBe(false)
    expect(hasText('chevron')).toBe(false)
  })

  it('SHAPE_CATEGORIES accounts for every kind exactly once — no gaps, no dupes', () => {
    const seen = new Map<string, number>()
    for (const cat of SHAPE_CATEGORIES) {
      for (const k of cat.kinds) seen.set(k, (seen.get(k) ?? 0) + 1)
    }
    expect(seen.size).toBe(SHAPE_KINDS.length)
    for (const k of SHAPE_KINDS) expect(seen.get(k)).toBe(1)
  })

  it('reports browser-preview fidelity instead of presenting fallbacks as exact geometry', () => {
    expect(shapePreviewFidelity('ellipse')).toBe('distinct')
    expect(shapePreviewFidelity('bentArrow')).toBe('approximated')
    expect(shapePreviewFidelity('heart')).toBe('generic')

    const counts = { distinct: 0, approximated: 0, generic: 0 }
    for (const kind of SHAPE_KINDS) counts[shapePreviewFidelity(kind)] += 1
    expect(counts.distinct).toBeGreaterThan(80)
    expect(counts.approximated).toBeGreaterThan(0)
    expect(counts.generic).toBeLessThan(20)
  })
})
