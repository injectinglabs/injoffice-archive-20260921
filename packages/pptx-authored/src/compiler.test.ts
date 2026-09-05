import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { NativeFontManifest, NativeFontResolver, NativeTextShaper } from '@injoffice/font-metrics/layout'
import { validateNativePptx } from '@injoffice/pptx-native'
import { compileNativePptxSlide, createRecordingPaintSurface, paintSlideRenderTree } from '@injoffice/pptx-render'
import { compileDeckSpecToNativeV1, compileWireDeckToNativeV1 } from './index.js'

const rectDeck = {
  slides: [{
    background: '#FFFFFF',
    shapes: [
      { kind: 'rect', key: 'back', x: 10, y: 20, cx: 300, cy: 200, fill: '#112233' },
      { kind: 'ellipse', key: 'front', x: 40, y: 50, cx: 100, cy: 80, fill: '#AABBCC' },
    ],
  }],
} as const

describe('native authored deck compiler', () => {
  it('is deterministic, validates its native output, and carries exact authored authority', () => {
    const first = compileWireDeckToNativeV1(rectDeck)
    const second = compileWireDeckToNativeV1(structuredClone(rectDeck))
    expect(first).toEqual(second)
    expect(first.ok).toBe(true)
    if (!first.ok) return
    expect(validateNativePptx(first.deck)).toEqual({ ok: true, value: first.deck })
    expect(first.deck).toMatchObject({
      contractVersion: 'pptx-native/v1',
      origin: 'authored',
      size: { cx: 12_192_000, cy: 6_858_000 },
      assets: [],
      compatibility: { status: 'editable', diagnostics: [] },
      slides: [{ provenance: 'authored', background: 'FFFFFF', compatibility: { status: 'editable', diagnostics: [] } }],
    })
    expect(first.deck.slides[0]!.elements.map((element) => element.provenance)).toEqual(['authored', 'authored'])
    expect(Object.isFrozen(first.deck)).toBe(true)
  })

  it('materializes the exact authored text, animation, transition, and no-paint defaults', () => {
    const result = compileWireDeckToNativeV1({
      slides: [{
        background: '#010203', transition: 'push',
        shapes: [{
          kind: 'textBox', placeholder: 'title', name: '', x: 100, y: 200, cx: 2_000_000, cy: 1_000_000,
          fill: '', stroke: '',
          paragraphs: [{ runs: [{ text: 'Exact', sizePt: 12.5, color: '#aabbcc', font: 'Aptos', bold: true }], align: 'ctr' }],
          enter: 'flyIn', enterDirection: 'bottom', enterDurationMs: 0, enterDistance: 0,
        }],
      }],
    })
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.deck.slides[0]).toMatchObject({ transition: { type: 'push', direction: 'left' } })
    expect(result.deck.slides[0]!.elements[0]).toMatchObject({
      kind: 'text', name: 'textBox', placeholder: 'title',
      animation: { effect: 'flyIn', direction: 'down', delayMs: 0, durationMs: 500, distancePpm: 250_000 },
      paragraphs: [{ align: 'center', level: 0, bullet: false, runs: [{ text: 'Exact', bold: true, fontSizeHundredthPt: 1250, color: 'AABBCC', fontFamily: 'Aptos' }] }],
      textBody: {
        leftInsetEmu: 91_440, rightInsetEmu: 91_440, topInsetEmu: 45_720, bottomInsetEmu: 45_720,
        wrap: 'square', verticalAnchor: 'top', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow',
      },
    })
    expect(result.deck.slides[0]!.elements[0]).not.toHaveProperty('fill')
    expect(result.deck.slides[0]!.elements[0]).not.toHaveProperty('stroke')
  })

  it('uses durable stable keys and collision-safe occurrence IDs without array-index fallback', () => {
    const original = compileWireDeckToNativeV1(rectDeck)
    const reordered = compileWireDeckToNativeV1({ slides: [{ background: '#FFFFFF', shapes: [...rectDeck.slides[0].shapes].reverse() }] })
    expect(original.ok && reordered.ok).toBe(true)
    if (!original.ok || !reordered.ok) return
    const byName = (deck: typeof original.deck) => new Map(deck.slides[0]!.elements.map((element) => [element.name, element.id]))
    expect(byName(original.deck)).toEqual(byName(reordered.deck))

    const duplicate = compileWireDeckToNativeV1({ slides: [{ background: '#FFFFFF', shapes: [
      { kind: 'rect', x: 1, y: 1, cx: 10, cy: 10, fill: '#000000' },
      { kind: 'rect', x: 1, y: 1, cx: 10, cy: 10, fill: '#000000' },
    ] }] })
    expect(duplicate.ok).toBe(true)
    if (duplicate.ok) expect(new Set(duplicate.deck.slides[0]!.elements.map((element) => element.id)).size).toBe(2)

    const duplicateKey = compileWireDeckToNativeV1({ slides: [{ background: '#FFFFFF', shapes: [
      { kind: 'rect', key: 'same', x: 1, y: 1, cx: 10, cy: 10 },
      { kind: 'ellipse', key: 'same', x: 20, y: 1, cx: 10, cy: 10 },
    ] }] })
    expect(duplicateKey).toMatchObject({ ok: false, issues: [{ code: 'authored.identifier' }] })
  })

  it('fails atomically for unsupported, inherited, lossy, and malformed WireDeck semantics', () => {
    const cases: Array<[unknown, string]> = [
      [{ slides: [{ background: '#fff', shapes: [] }] }, 'authored.color'],
      [{ slides: [{ shapes: [] }] }, 'authored.inheritedBackground'],
      [{ slides: [{ background: '#FFFFFF', shapes: [{ kind: 'line', x: 0, y: 0, cx: 100, cy: 0, stroke: '#000000' }] }] }, 'authored.integerEmu'],
      [{ slides: [{ background: '#FFFFFF', shapes: [{ kind: 'textBox', x: 0, y: 0, cx: 1_000_000, cy: 500_000, paragraphs: [{ runs: [{ text: 'inherits' }] }] }] }] }, 'authored.inheritedTextStyle'],
      [{ slides: [{ background: '#FFFFFF', shapes: [{ kind: 'textBox', x: 0, y: 0, cx: 1_000_000, cy: 500_000, fill: '#FFFFFF', paragraphs: [] }] }] }, 'authored.textBoxPaint'],
      [{ slides: [{ background: '#FFFFFF', shapes: [], html: '<b>authority</b>' }] }, 'authored.unsupportedField'],
    ]
    for (const [input, code] of cases) {
      const result = compileWireDeckToNativeV1(input)
      expect(result.ok, code).toBe(false)
      if (!result.ok) expect(result.issues.map((issue) => issue.code), code).toContain(code)
      expect(result).not.toHaveProperty('deck')
    }
  })

  it('compiles DeckSpec through stable authored identities and refuses silently dropped authoring fields', () => {
    const spec = {
      id: 'sales-deck', title: 'Sales', theme: 'boardroom', slides: [
        { id: 'intro', kind: 'title', title: 'Native Office', subtitle: 'No browser layout' },
        { id: 'facts', kind: 'bullets', title: 'Facts', bullets: ['One', '**Two**'] },
      ],
    }
    const first = compileDeckSpecToNativeV1(spec)
    const changedText = compileDeckSpecToNativeV1({ ...spec, title: 'Renamed', slides: [{ ...spec.slides[0], title: 'Edited' }, spec.slides[1]] })
    expect(first.ok && changedText.ok).toBe(true)
    if (first.ok && changedText.ok) {
      expect(first.deck.documentId).toBe(changedText.deck.documentId)
      expect(first.deck.slides.map((slide) => slide.id)).toEqual(changedText.deck.slides.map((slide) => slide.id))
      expect(first.deck.slides[0]!.elements.map((element) => element.id)).toEqual(changedText.deck.slides[0]!.elements.map((element) => element.id))
    }

    const notes = compileDeckSpecToNativeV1({ ...spec, slides: [{ ...spec.slides[0], notes: 'not in v1' }] })
    expect(notes).toMatchObject({ ok: false, issues: [{ code: 'authored.unsupportedNotes' }] })
    const stale = compileDeckSpecToNativeV1({ ...spec, slides: [{ ...spec.slides[0], shapeOverrides: { missing: { x: 10 } } }] })
    expect(stale).toMatchObject({ ok: false, issues: [{ code: 'authored.staleShapeKey' }] })
    const unknownTheme = compileDeckSpecToNativeV1({ ...spec, theme: 'silent-fallback' })
    expect(unknownTheme).toMatchObject({ ok: false, issues: [{ code: 'authored.theme' }] })
  })

  it('feeds the native RenderTree in authored z-order and paints without a DOM authority', async () => {
    const result = compileWireDeckToNativeV1(rectDeck)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    const tree = await compileNativePptxSlide(result.deck, 0, { textLayout })
    expect(tree.nodes.map((node) => [node.zIndex, node.kind, node.sourceElementId])).toEqual([
      [0, 'shape', result.deck.slides[0]!.elements[0]!.id],
      [1, 'shape', result.deck.slides[0]!.elements[1]!.id],
    ])
    const surface = createRecordingPaintSurface()
    await paintSlideRenderTree(tree, surface)
    expect(surface.commands.some((command) => command.kind === 'path' && command.fill !== undefined)).toBe(true)
  })
})

const digest = 'sha256:054edec1d0211f624fed0cbca9d4f9400b0e491c43742af2c5b0abebf0c990d8' as const
const manifest: NativeFontManifest = {
  version: 1,
  manifestId: 'unused-shape-only-fixture',
  revision: '1',
  faces: [{
    faceId: 'fixture.regular', family: 'Fixture Sans', weight: 400, style: 'normal', stretch: 100,
    source: { kind: 'bundled', resourceId: 'fixture-font', contentDigest: digest },
  }],
  fallbackChains: [{ chainId: 'fixture.default', faceIds: ['fixture.regular'] }],
}
const resolver: NativeFontResolver = {
  providerId: 'unused-resolver', providerRevision: '1',
  resolve() { throw new Error('shape-only fixture must not resolve fonts') },
  load() { throw new Error('shape-only fixture must not load fonts') },
}
const shaper: NativeTextShaper = {
  providerId: 'unused-shaper', providerRevision: '1',
  shape() { throw new Error('shape-only fixture must not shape text') },
}
const textLayout = {
  manifest,
  resolver,
  shaper,
  defaults: { fontFamilies: ['Fixture Sans'], fontSizeHundredthPt: 1_000, script: 'Latn', language: 'en-US', direction: 'ltr' as const, fallbackChainIds: ['fixture.default'] },
}

describe('architecture boundary', () => {
  it('keeps the production compiler and authoring entry free of legacy/browser authority imports', () => {
    const root = resolve(import.meta.dirname, '../../..')
    const compiler = readFileSync(resolve(import.meta.dirname, 'compiler.ts'), 'utf8')
    const pureAuthoringFiles = ['authoring.ts', 'compile.ts', 'diagram.ts', 'themes.ts', 'types.ts', 'wire.ts']
      .map((file) => readFileSync(resolve(root, `packages/slides/src/${file}`), 'utf8'))
    const authoring = pureAuthoringFiles.join('\n')
    const slidesManifest = JSON.parse(readFileSync(resolve(root, 'packages/slides/package.json'), 'utf8'))
    const imports = [...`${compiler}\n${authoring}`.matchAll(/from\s+['"]([^'"]+)['"]/g)].map((match) => match[1])
    expect(imports).not.toEqual(expect.arrayContaining(['react', 'react-dom', 'react-konva', 'konva', '@injoffice/slides']))
    expect(imports.some((specifier) => /DeckView|DeckCanvasView/i.test(specifier))).toBe(false)
    expect(compiler).not.toMatch(/\b(?:window|HTMLElement|ResizeObserver|OffscreenCanvas)\b|\bdocument\.|getBoundingClientRect|measureText|createElement/)
    const authoringExternalImports = [...authoring.matchAll(/from\s+['"]([^.'"][^'"]*)['"]/g)].map((match) => match[1])
    expect(authoringExternalImports).toEqual([])
    expect(slidesManifest.exports['./authoring']).toEqual({ types: './dist/authoring.d.ts', import: './dist/authoring.js' })
    for (const file of ['DeckView.tsx', 'DeckCanvasView.tsx']) {
      expect(readFileSync(resolve(root, `packages/slides/src/${file}`), 'utf8')).toContain('@deprecated')
    }
  })
})
