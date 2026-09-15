import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { NativeEvaluatedGeometry, NativePptxDeck } from '@injoffice/pptx-native'
import { compileNativePptxSlide, createRecordingPaintSurface, paintSlideRenderTree, type NativePptxTextLayout, type RenderNode } from './index.js'
import { PICTURE_GEOMETRY_CLIP_POLICY } from './pictureGeometryClip.js'

const root = resolve(import.meta.dirname, '../../..')
const parsedFull = JSON.parse(readFileSync(resolve(root, 'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'), 'utf8')) as NativePptxDeck
/** Pictures need no shaping; a geometry-only layout keeps the slide free of font fixtures. */
const textLayout = (): NativePptxTextLayout => ({
  manifest: { version: 1, manifestId: 'geometry-only', revision: '1', faces: [{ faceId: 'unused', family: 'Unused', weight: 400, style: 'normal', stretch: 100, source: { kind: 'host', resourceId: 'unused', contentDigest: `sha256:${'0'.repeat(64)}` } }], fallbackChains: [] },
  resolver: { providerId: 'unused', providerRevision: '1', resolve() { throw Error('No text') }, load() { throw Error('No font') } },
  shaper: { providerId: 'unused', providerRevision: '1', shape() { throw Error('No glyphs') } },
  defaults: { fontFamilies: ['Unused'], fontSizeHundredthPt: 1200, script: 'Latn', language: 'en-US', direction: 'ltr' },
})

/** An ellipse outline the way the Go catalog evaluator emits it for a 3,000,000 × 2,000,000 frame. */
function ellipse(fillMode: NativeEvaluatedGeometry['paths'][number]['fillMode'] = 'norm'): NativeEvaluatedGeometry {
  const arc = (x: number, y: number) => ({ kind: 'arcTo' as const, x, y, rx: 1_500_000, ry: 1_000_000, largeArc: false, clockwise: true })
  return { profile: 'drawingml-paths-v1', textRect: { x: 439_340, y: 292_893, cx: 2_121_320, cy: 1_414_214 }, paths: [{ fillMode, stroke: true, commands: [{ kind: 'moveTo', x: 0, y: 1_000_000 }, arc(1_500_000, 0), arc(3_000_000, 1_000_000), arc(1_500_000, 2_000_000), arc(0, 1_000_000), { kind: 'close' }] }] }
}

function pictureDeck() {
  const deck = structuredClone(parsedFull)
  const picture = deck.slides[0]!.elements.find((element) => element.kind === 'picture')!
  if (picture.kind !== 'picture') throw new Error('picture fixture')
  deck.slides[0]!.elements = [picture]
  picture.transform.cx = 3_000_000; picture.transform.cy = 2_000_000
  picture.crop = { left: 10000, top: 20000, right: 30000, bottom: 0 }
  picture.geometry = ellipse()
  picture.compatibility = { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: PICTURE_GEOMETRY_CLIP_POLICY, message: 'catalog outline evaluated from source' }] }
  return { deck, picture }
}

function findNode<T extends RenderNode['kind']>(tree: Awaited<ReturnType<typeof compileNativePptxSlide>>, kind: T, id: string): Extract<RenderNode, { kind: T }> {
  const node = tree.nodes.find((candidate) => candidate.kind === kind && candidate.sourceElementId === id)
  if (!node) throw new Error(`missing ${kind} ${id}`)
  return node as Extract<RenderNode, { kind: T }>
}

describe('picture preset-catalog clip', () => {
  it('clips the image to the evaluated outline in its local frame and keeps the source crop', async () => {
    const { deck, picture } = pictureDeck()
    const before = JSON.stringify(deck)
    const tree = await compileNativePptxSlide(deck, 0, { textLayout: textLayout() })
    const image = findNode(tree, 'image', picture.id)
    expect(image.clip).toEqual({ kind: 'path', rect: { x: 0, y: 0, cx: 3_000_000, cy: 2_000_000 }, path: ellipse().paths[0]!.commands })
    expect(image.crop).toEqual(picture.crop)
    expect(tree.diagnostics).toContainEqual(expect.objectContaining({ severity: 'warning', code: 'picture.presetCatalogClipPreview', elementId: picture.id, message: expect.stringContaining(PICTURE_GEOMETRY_CLIP_POLICY) }))
    const surface = createRecordingPaintSurface(); paintSlideRenderTree(tree, surface)
    const commands = surface.finish(), index = commands.findIndex((command) => command.kind === 'image' && command.sourceElementId === picture.id)
    expect(commands[index - 1]).toEqual({ kind: 'clipPath', rect: image.bounds, path: ellipse().paths[0]!.commands })
    expect(commands[index - 2]).toEqual({ kind: 'transform', transform: image.transform })
    expect(JSON.stringify(deck)).toBe(before)
  })
  it('keeps the exact roundRect clip contract unchanged when no geometry is present', async () => {
    const { deck, picture } = pictureDeck()
    delete picture.geometry; picture.clip = 'roundRect'
    const image = findNode(await compileNativePptxSlide(deck, 0, { textLayout: textLayout() }), 'image', picture.id)
    expect(image.clip).toEqual({ kind: 'roundRect', rect: { x: 0, y: 0, cx: 3_000_000, cy: 2_000_000 }, radiusEmu: 333_340 })
  })
  it('still refuses catalog failures as a placeholder instead of guessing an outline', async () => {
    const { deck, picture } = pictureDeck()
    delete picture.geometry
    picture.compatibility.diagnostics = [{ severity: 'warning', code: 'pptx.picture-geometry-unavailable', message: 'outside the evaluated preset catalog' }]
    const tree = await compileNativePptxSlide(deck, 0, { textLayout: textLayout() })
    expect(findNode(tree, 'placeholder', picture.id).label).toBe('Unsupported picture geometry preserved')
  })
  it('rejects editable, clip-conflicting, stroke-only and malformed picture geometry', async () => {
    for (const [name, mutate] of [
      ['editable', (picture: ReturnType<typeof pictureDeck>['picture']) => { picture.compatibility.status = 'editable' }],
      ['clip conflict', (picture: ReturnType<typeof pictureDeck>['picture']) => { picture.clip = 'roundRect' }],
      ['stroke-only outline', (picture: ReturnType<typeof pictureDeck>['picture']) => { picture.geometry = ellipse('none') }],
      ['no paths', (picture: ReturnType<typeof pictureDeck>['picture']) => { picture.geometry = { ...ellipse(), paths: [] } }],
      ['unsafe coordinate', (picture: ReturnType<typeof pictureDeck>['picture']) => { picture.geometry = ellipse(); Object.assign(picture.geometry.paths[0]!.commands[0]!, { x: 2 ** 53 }) }],
    ] as const) {
      const { deck, picture } = pictureDeck()
      mutate(picture)
      await expect(compileNativePptxSlide(deck, 0, { textLayout: textLayout() }), name).rejects.toThrow()
    }
  })
})
