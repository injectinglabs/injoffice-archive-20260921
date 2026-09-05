import { readFileSync } from 'node:fs'
import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'
import { deckFromOutline, resetSlideIds } from '@injoffice/slides'
import { createPdfDemoFixture, PDF_DEMO_FILE_NAME, PDF_DEMO_PAGE_TITLES } from './pdfDemoFixture'
import {
  makeAuthoredPresentationDemo,
  PRESENTATION_DEMO_OUTLINE,
  PRESENTATION_DEMO_TITLE,
} from './presentationDemoFixtures'

describe('playground presentation fixtures', () => {
  it('compiles a credible multi-layout authored deck into the native contract', () => {
    const result = compileDeckSpecToNativeV1(makeAuthoredPresentationDemo(
      PRESENTATION_DEMO_TITLE,
      'Q3 operating brief · September 2026',
      'boardroom',
    ))

    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.deck.slides).toHaveLength(4)
    expect(result.deck.slides.every((slide) => slide.elements.length >= 2)).toBe(true)
    expect(result.deck.slides.map((slide) => slide.transition?.type)).toEqual(['fade', 'push', 'wipe', 'fade'])
    expect(result.deck.compatibility.status).toBe('editable')
  })

  it('starts the editable Slides workbench with a substantial outline', () => {
    resetSlideIds()
    const deck = deckFromOutline(PRESENTATION_DEMO_TITLE, PRESENTATION_DEMO_OUTLINE)

    expect(deck.slides).toHaveLength(6)
    expect(deck.slides.map((slide) => slide.kind)).toEqual([
      'title', 'bullets', 'bullets', 'bullets', 'quote', 'section',
    ])
    expect(deck.slides.flatMap((slide) => slide.bullets ?? [])).toHaveLength(10)
    expect(deck.slides[0]?.notes).toContain('customer outcome')
  })

  it('uses the compact, reproducible, populated PPTX proof fixture', () => {
    const source = readFileSync(new URL('./pages/PptxNativePage.tsx', import.meta.url), 'utf8')
    const fixture = readFileSync(new URL('../../../go/pptxpatch/testdata/playground_northstar_review.pptx', import.meta.url))

    expect(source).toContain('go/pptxpatch/testdata/playground_northstar_review.pptx')
    expect(source).toContain('populated, reproducible three-slide launch review')
    expect(fixture.subarray(0, 2).toString()).toBe('PK')
    expect(fixture.byteLength).toBeGreaterThan(10_000)
    expect(fixture.byteLength).toBeLessThan(100_000)
  })
})

describe('playground PDF fixture', () => {
  it('builds the same realistic four-page approval packet every time', async () => {
    const first = await createPdfDemoFixture()
    const second = await createPdfDemoFixture()
    expect(first).toEqual(second)

    const document = await PDFDocument.load(first)
    expect(PDF_DEMO_FILE_NAME).toBe('northstar-operating-review.pdf')
    expect(document.getPageCount()).toBe(PDF_DEMO_PAGE_TITLES.length)
    expect(document.getTitle()).toBe('Northstar operating review')
    expect(document.getAuthor()).toBe('InjOffice demo')
    expect(document.getCreationDate().toISOString()).toBe('2026-09-01T12:00:00.000Z')
    expect(document.getForm().getFields().map((field) => field.getName())).toEqual([
      'release.decisionOwner',
      'release.rolloutTier',
      'release.decisionNotes',
      'release.approved',
    ])
    expect(first.byteLength).toBeGreaterThan(10_000)
  })
})
