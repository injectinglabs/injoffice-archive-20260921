import { describe, expect, it } from 'vitest'
import type { NativeFontManifest, NativeFontResolver, NativeTextShaper } from '@injoffice/font-metrics/layout'
import type { NativeElement, NativePptxDeck } from '@injoffice/pptx-native'
import { compileNativePptxSlide, createRecordingPaintSurface, paintSlideRenderTree, type NativePptxTextLayout } from './index.js'

const digest = `sha256:${'b'.repeat(64)}` as const
const manifest: NativeFontManifest = {
  version: 1,
  manifestId: 'group-adversarial',
  revision: '1',
  faces: [{
    faceId: 'group.regular', family: 'Group Fixture', aliases: ['Group Fixture'], weight: 400, style: 'normal', stretch: 100,
    source: { kind: 'bundled', resourceId: 'group-font', contentDigest: digest },
  }],
  fallbackChains: [{ chainId: 'group.default', faceIds: ['group.regular'] }],
}
const resolver: NativeFontResolver = {
  providerId: 'unused-group-resolver', providerRevision: '1',
  resolve() { throw new Error('shape-only group test must not resolve text') },
  load() { throw new Error('shape-only group test must not load fonts') },
}
const shaper: NativeTextShaper = {
  providerId: 'unused-group-shaper', providerRevision: '1',
  shape() { throw new Error('shape-only group test must not shape text') },
}
const textLayout: NativePptxTextLayout = {
  manifest, resolver, shaper,
  defaults: {
    fontFamilies: ['Group Fixture'], fontSizeHundredthPt: 1_000,
    script: 'Latn', language: 'en-US', direction: 'ltr', fallbackChainIds: ['group.default'],
  },
}

function deckWith(element: NativeElement): NativePptxDeck {
  return {
    contractVersion: 'pptx-native/v1', documentId: 'group-adversarial-deck', origin: 'authored',
    size: { cx: 1_000, cy: 1_000 }, assets: [],
    slides: [{
      id: 'group-adversarial-slide', provenance: 'authored', elements: [element], passthrough: [],
      compatibility: { status: 'editable', diagnostics: [] },
    }],
    compatibility: { status: 'editable', diagnostics: [] },
  }
}

describe('native group projection adversarial boundaries', () => {
  it('keeps asymmetric axes and negative child coordinates distinct through paint', async () => {
    const leaf: NativeElement = {
      kind: 'shape', id: 'axis-leaf', provenance: 'authored', preset: 'rect', paragraphs: [],
      transform: { x: -11, y: 21, cx: 2, cy: 3 }, stroke: { color: '000000', widthEmu: 7 },
      passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }
    const group: NativeElement = {
      kind: 'group', id: 'axis-group', provenance: 'authored',
      transform: { x: -100, y: 200, cx: 6, cy: 15 },
      childTransform: { x: -10, y: 20, cx: 2, cy: 3 },
      children: [leaf], passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }

    const tree = await compileNativePptxSlide(deckWith(group), 0, { textLayout })
    expect(tree.nodes[0]).toMatchObject({
      kind: 'group',
      transform: { aPpm: 3_000_000, dPpm: 5_000_000, txEmu: -70, tyEmu: 100 },
      bounds: { x: -10, y: 20, cx: 2, cy: 3 },
    })

    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(tree, recording)
    const transforms = recording.finish().filter((command) => command.kind === 'transform')
    expect(transforms).toEqual([
      { kind: 'transform', transform: { aPpm: 3_000_000, bPpm: 0, cPpm: 0, dPpm: 5_000_000, txEmu: -70, tyEmu: 100 } },
      { kind: 'transform', transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: -11, tyEmu: 21 } },
    ])

    // Applying the emitted transforms puts the leaf origin at (-103, 205).
    // This catches swapped axes and incorrect chOff translation signs.
    expect(-70 + 3 * -11).toBe(-103)
    expect(100 + 5 * 21).toBe(205)
  })

  it('inherits the exact group affine for straight connectors without clipping their centered stroke', async () => {
    const connector: NativeElement = {
      kind: 'connector', id: 'grouped-connector', provenance: 'authored',
      transform: { x: 12, y: 25, cx: 30, cy: 40 }, flipH: true,
      stroke: { color: '123456', widthEmu: 7, cap: 'flat', join: 'round', dash: 'solid' },
      passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }
    const group: NativeElement = {
      kind: 'group', id: 'connector-group', provenance: 'authored',
      transform: { x: 100, y: 200, cx: 60, cy: 90 },
      childTransform: { x: 10, y: 20, cx: 30, cy: 30 },
      children: [connector], passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }

    const tree = await compileNativePptxSlide(deckWith(group), 0, { textLayout })
    expect(tree.nodes[0]).toMatchObject({
      kind: 'group',
      transform: { aPpm: 2_000_000, dPpm: 3_000_000, txEmu: 80, tyEmu: 140 },
      children: [{
        kind: 'connector',
        transform: { txEmu: 12, tyEmu: 25 },
        bounds: { x: 0, y: 0, cx: 30, cy: 40 },
        path: [{ kind: 'moveTo', x: 30, y: 0 }, { kind: 'lineTo', x: 0, y: 40 }],
        stroke: { color: '123456', widthEmu: 7, cap: 'flat', join: 'round', dash: 'solid' },
      }],
    })
    const groupNode = tree.nodes[0]!
    expect(groupNode).not.toHaveProperty('clip')
    if (groupNode.kind !== 'group') throw new Error('expected grouped connector fixture')
    expect(groupNode.children[0]).not.toHaveProperty('clip')

    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(tree, recording)
    const commands = recording.finish()
    expect(commands.filter((command) => command.kind === 'clipRect')).toEqual([
      { kind: 'clipRect', rect: { x: 0, y: 0, cx: 1_000, cy: 1_000 } },
    ])
    expect(commands).toContainEqual({
      kind: 'path', sourceElementId: 'grouped-connector',
      path: [{ kind: 'moveTo', x: 30, y: 0 }, { kind: 'lineTo', x: 0, y: 40 }],
      stroke: { color: '123456', widthEmu: 7, cap: 'flat', join: 'round', dash: 'solid' },
      headArrow: false, tailArrow: false,
    })
  })
})
