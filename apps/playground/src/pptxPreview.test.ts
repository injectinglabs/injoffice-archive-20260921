import { describe, expect, it } from 'vitest'
import type { NativeAsset, NativeElement, NativePptxDeck } from '@injoffice/pptx-native'
import { previewColor, previewImage, previewIssue, previewSlideSize } from './pptxPreview'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import PptxFilePreview from './components/PptxFilePreview'

const asset = (overrides: Partial<NativeAsset> = {}): NativeAsset => ({ id: 'image', provenance: 'parsed', contentType: 'image/png', sha256: 'a'.repeat(64), byteLength: 3, dataBase64: 'YWJj', passthrough: [], ...overrides })
const element = (overrides: object): NativeElement => ({ kind: 'picture', id: 'picture', provenance: 'parsed', assetId: 'image', transform: { x: 0, y: 0, cx: 100, cy: 100 }, compatibility: { status: 'editable', diagnostics: [] }, passthrough: [], ...overrides } as NativeElement)

describe('PPTX approximate preview policy', () => {
  it('accepts only literal RGB colors, not document-controlled CSS', () => {
    expect(previewColor('aBcDeF')).toBe('#aBcDeF')
    expect(previewColor('url(https://example.com/image)')).toBe('#ffffff')
    expect(previewColor(undefined, 'none')).toBe('none')
  })
  it('embeds bounded raster bytes and refuses active or remote resources', () => {
    expect(previewImage(asset())).toBe('data:image/png;base64,YWJj')
    expect(previewImage(asset({ contentType: 'image/jpeg' }))).toBe('data:image/jpeg;base64,YWJj')
    expect(previewImage(asset({ contentType: 'image/svg+xml' }))).toBeUndefined()
    expect(previewImage(asset({ dataBase64: 'https://example.com/image' }))).toBeUndefined()
    expect(previewImage(asset({ dataBase64: 'a'.repeat(4_000_001) }))).toBeUndefined()
    expect(previewImage(undefined)).toBeUndefined()
  })
  it('keeps missing content explicit and never relaxes refused object policy', () => {
    expect(previewIssue(element({}), [asset()])).toBeUndefined()
    expect(previewIssue(element({ compatibility: { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: 'pptx.picture-crop-unavailable', message: 'outset crop' }] } }), [asset()])).toContain('crop cannot be previewed')
    expect(previewIssue(element({}), [])).toContain('unavailable')
    expect(previewIssue(element({ kind: 'chart', chart: {} }), [])).toContain('no embedded preview')
    expect(previewIssue(element({ kind: 'group', children: [] }), [])).toContain('Grouped content')
    expect(previewIssue(element({ compatibility: { status: 'refused', diagnostics: [] } }), [asset()])).toContain('Unsupported')
  })
  it('preserves slide aspect ratio and rejects invalid or excessive dimensions', () => {
    const deck = (cx: number, cy: number) => ({ size: { cx, cy } } as NativePptxDeck)
    expect(previewSlideSize(deck(12_192_000, 6_858_000))).toMatchObject({ width: 960, height: 540 })
    for (const [cx, cy] of [[0, 1], [Infinity, 1], [1, NaN], [1, 1000], [-1, 1]]) expect(previewSlideSize(deck(cx, cy))).toBeUndefined()
  })
  it('renders a real-file projection with a persistent fidelity label and missing-content list', () => {
    const deck: NativePptxDeck = {
      contractVersion: 'pptx-native/v1', documentId: 'preview-test', origin: 'parsed',
      size: { cx: 12_192_000, cy: 6_858_000 }, assets: [], compatibility: { status: 'editable', diagnostics: [] },
      slides: [{ id: 'slide', provenance: 'parsed', passthrough: [], compatibility: { status: 'editable', diagnostics: [] }, elements: [
        element({ kind: 'text', textBody: { leftInsetEmu: 0, rightInsetEmu: 0, topInsetEmu: 0, bottomInsetEmu: 0, wrap: 'none', verticalAnchor: 'top', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow' }, paragraphs: [{ runs: [{ text: '<script>alert(1)</script>', bold: true, fontSizeHundredthPt: 2400 }] }] }),
        element({ id: 'chart', kind: 'chart', chart: {} }),
        element({ id: 'invalid', transform: { x: 0, y: 0, cx: 0, cy: 100 } }),
      ] }],
    }
    const before = JSON.stringify(deck)
    const html = renderToStaticMarkup(createElement(PptxFilePreview, { deck }))
    expect(html).toContain('Approximate file preview')
    expect(html).toContain('Not PowerPoint-equivalent rendering')
    expect(html).toContain('viewBox="0 0 960 540"')
    expect(html).toContain('no embedded preview image')
    expect(html).toContain('preview geometry unavailable')
    expect(html).toContain('&lt;script&gt;')
    expect(html).toContain('white-space:pre;')
    expect(html).toContain('overflow:visible')
    expect(html).not.toContain('overflow-wrap:anywhere')
    expect(html).not.toContain('<script>')
    expect(JSON.stringify(deck)).toBe(before)
  })
  it('crops the original raster through a bounded SVG source viewport without changing bytes', () => {
    const deck: NativePptxDeck = {
      contractVersion: 'pptx-native/v1', documentId: 'crop-preview', origin: 'authored',
      size: { cx: 100, cy: 100 }, assets: [asset()], compatibility: { status: 'editable', diagnostics: [] },
      slides: [{ id: 'slide', provenance: 'authored', passthrough: [], compatibility: { status: 'editable', diagnostics: [] }, elements: [element({ crop: { left: 12500, top: 25000, right: 37500, bottom: 0 } })] }],
    }
    const before = JSON.stringify(deck)
    const html = renderToStaticMarkup(createElement(PptxFilePreview, { deck }))
    expect(html).toContain('viewBox="12500 25000 50000 75000"')
    expect(html).toContain('overflow="hidden"')
    expect(html).toContain('width="100000" height="100000" preserveAspectRatio="none"')
    expect(html).toContain('href="data:image/png;base64,YWJj"')
    expect(JSON.stringify(deck)).toBe(before)
  })
})
