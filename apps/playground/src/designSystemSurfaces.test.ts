import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'

const surfaces = readFileSync(new URL('./design-system/surfaces.tsx', import.meta.url), 'utf8')
const gallery = readFileSync(new URL('./design-system/GalleryPage.tsx', import.meta.url), 'utf8')
const workbench = readFileSync(new URL('./workbench.css', import.meta.url), 'utf8')

describe('design-system surface mocks', () => {
  it('covers overview, every demo surface, and guides', () => {
    const expected = ['overview', ...DEMOS.map((demo) => demo.surface), 'guides']
    for (const surface of expected) {
      expect(surfaces, surface).toContain(`id: 'surface-${surface}'`)
    }
  })

  it('keeps the gallery independent of Univer Pro and DeckView', () => {
    expect(gallery).toContain('SURFACE_GALLERY')
    expect(surfaces).not.toContain('@univerjs-pro')
    expect(surfaces).not.toContain('@univerjs/presets')
    expect(surfaces).not.toContain('DeckView')
    expect(workbench).not.toContain('.ds-sheet-frame')
  })
})
