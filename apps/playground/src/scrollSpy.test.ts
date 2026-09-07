import { describe, expect, it } from 'vitest'
import { hashForSurface, pickActiveSurface, prefersReducedMotion, surfaceSectionId, viewportProbeY, type SectionMetrics } from './scrollSpy'

const sections: SectionMetrics[] = [
  { surface: 'overview', top: 0, bottom: 800 },
  { surface: 'sheets', top: 800, bottom: 1600 },
  { surface: 'docs', top: 1600, bottom: 2400 },
]

describe('playground scroll spy', () => {
  it('names demo sections so hash routes can scroll to them', () => {
    expect(surfaceSectionId('overview')).toBe('demo-overview')
    expect(surfaceSectionId('sheets')).toBe('demo-sheets')
    expect(surfaceSectionId('pptx-native')).toBe('demo-pptx-native')
  })

  it('selects the last section that has crossed the probe line', () => {
    expect(pickActiveSurface(sections, 100)).toBe('overview')
    expect(pickActiveSurface(sections, 800)).toBe('sheets')
    expect(pickActiveSurface(sections, 1599)).toBe('sheets')
    expect(pickActiveSurface(sections, 1600)).toBe('docs')
    expect(pickActiveSurface(sections, 3000)).toBe('docs')
  })

  it('falls back to the first section when every block is still below the probe', () => {
    const below = sections.map((section) => ({ ...section, top: section.top + 400, bottom: section.bottom + 400 }))
    expect(pickActiveSurface(below, 100)).toBe('overview')
    expect(pickActiveSurface([], 0)).toBeUndefined()
  })

  it('preserves query strings while the surface stays the same', () => {
    expect(hashForSurface('sheets', '#/sheets?view=native')).toBe('#/sheets?view=native')
    expect(hashForSurface('sheets', '#/sheets')).toBe('#/sheets')
    expect(hashForSurface('docs', '#/sheets?view=native')).toBe('#/docs')
    expect(hashForSurface('overview', '')).toBe('#/overview')
  })

  it('reads reduced-motion from the media query', () => {
    expect(prefersReducedMotion({ matches: true })).toBe(true)
    expect(prefersReducedMotion({ matches: false })).toBe(false)
  })

  it('places the spy probe below the workbench header', () => {
    expect(viewportProbeY(800)).toBe(Math.round(800 * 0.22))
    expect(viewportProbeY(200)).toBe(88)
  })
})
