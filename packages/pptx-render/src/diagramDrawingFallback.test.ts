import { expect, it } from 'vitest'
import type { NativeFontManifest, NativeFontResolver, NativeTextShaper, ResolvedFontFace, ShapedCluster, ShapedGlyph } from '@injoffice/font-metrics/layout'
import { validateNativePptx, type NativeElement, type NativePptxDeck } from '@injoffice/pptx-native'
import { compileNativePptxSlide, createRecordingPaintSurface, paintSlideRenderTree, type NativePptxTextLayout, type PaintCommand } from './index.js'

// SmartArt drawing fallback (diagram-drawing-fallback-v1): the Go extractor
// projects PowerPoint's pre-laid-out dsp:sp shapes as a preserve-only group of
// evaluated-geometry shapes anchored at the graphic frame. These tests pin the
// render contract for that projection: ordinary paint, exact frame-relative
// placement, read-only labeling, and fail-closed validation.
const DIAGRAM_CODE = 'pptx.diagram-drawing-fallback-preview'
const FRAME = { x: 1_524_000, y: 1_397_000, cx: 6_096_000, cy: 4_064_000 }
const SLIDE_PART = 'ppt/slides/slide1.xml'
const digest = 'sha256:054edec1d0211f624fed0cbca9d4f9400b0e491c43742af2c5b0abebf0c990d8' as const
const manifest: NativeFontManifest = {
  version: 1, manifestId: 'diagram-fixture', revision: '1',
  faces: [{ faceId: 'fixture.regular', family: 'Fixture Sans', aliases: ['Calibri'], weight: 400, style: 'normal', stretch: 100, source: { kind: 'bundled', resourceId: 'fixture-font', contentDigest: digest } }],
  fallbackChains: [{ chainId: 'fixture.default', faceIds: ['fixture.regular'] }],
}
const face: ResolvedFontFace = { faceId: 'fixture.regular', family: 'Fixture Sans', weight: 400, style: 'normal', stretch: 100, sourceKind: 'bundled', resourceId: 'fixture-font', contentDigest: digest, resolution: 'substitute', matchedFamily: 'Calibri', fallbackChainId: 'fixture.default' }
const resolver: NativeFontResolver = {
  providerId: 'fixture-resolver', providerRevision: '1',
  resolve() { return { status: 'resolved', face, attemptedFaceIds: ['fixture.regular'], decisions: [] } },
  load(resolved) { return { face: resolved, bytes: new Uint8Array([0, 1, 2, 3]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 } } },
}
const shaper: NativeTextShaper = {
  providerId: 'fixture-shaper', providerRevision: '1',
  shape({ run, startUtf16, endUtf16, font }) {
    const glyphs: ShapedGlyph[] = []
    const clusters: ShapedCluster[] = []
    let utf16 = startUtf16
    for (const character of run.text.slice(startUtf16, endUtf16)) {
      const end = utf16 + character.length
      glyphs.push({ glyphId: character.codePointAt(0)!, clusterIndex: clusters.length, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 })
      clusters.push({ startUtf16: utf16, endUtf16: end, glyphStart: glyphs.length - 1, glyphEnd: glyphs.length, advanceInlineMilliPoints: 1_000, whitespace: /^\s$/u.test(character) })
      utf16 = end
    }
    const ascent = run.fontSizeMilliPoints * 8 / 10, descent = -run.fontSizeMilliPoints * 2 / 10, lineGap = run.fontSizeMilliPoints * 2 / 10
    return { startUtf16, endUtf16, face: font.face, glyphs, clusters, metrics: { fontSizeMilliPoints: run.fontSizeMilliPoints, ascentMilliPoints: ascent, descentMilliPoints: descent, lineGapMilliPoints: lineGap, lineHeightMilliPoints: ascent - descent + lineGap }, advanceInlineMilliPoints: clusters.length * 1_000, advanceBlockMilliPoints: 0 }
  },
}
const textLayout: NativePptxTextLayout = { manifest, resolver, shaper, defaults: { fontFamilies: ['Fixture Sans'], fontSizeHundredthPt: 1_000, script: 'Latn', language: 'en-US', direction: 'ltr', fallbackChainIds: ['fixture.default'] } }

function diagnostic(elementId: string, partName = 'ppt/diagrams/drawing1.xml') {
  return { severity: 'warning' as const, code: DIAGRAM_CODE, message: 'diagram shape painted verbatim from the pre-laid-out drawing part (diagram-drawing-fallback-v1); target remains read-only', scope: { slideId: 'slide-1', elementId, partName } }
}

/** One dsp:sp as the extractor emits it: frame-relative xfrm, evaluated rect paths, solid theme paint, style-resolved white text. */
function diagramShape(index: number, text: string, x: number, y: number): NativeElement {
  const id = `diagram-shape-${index}`
  const cx = 1_828_800, cy = 914_400
  return {
    kind: 'shape', id, provenance: 'parsed', transform: { x, y, cx, cy },
    geometry: { profile: 'drawingml-paths-v1', textRect: { x: 0, y: 0, cx, cy }, paths: [{ fillMode: 'norm', stroke: true, commands: [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'lineTo', x: cx, y: 0 }, { kind: 'lineTo', x: cx, y: cy }, { kind: 'lineTo', x: 0, y: cy }, { kind: 'close' }] }] },
    fill: '2F6FED', stroke: { color: 'FFFFFF', widthEmu: 12_700 },
    paragraphs: [{ align: 'center', level: 0, bullet: false, runs: [{ text, bold: false, italic: false, fontSizeHundredthPt: 2_000, color: 'FFFFFF', fontFamily: 'Calibri' }] }],
    textBody: { leftInsetEmu: 36_195, rightInsetEmu: 36_195, topInsetEmu: 36_195, bottomInsetEmu: 36_195, wrap: 'square', verticalAnchor: 'center', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow' },
    source: { partName: SLIDE_PART, objectId: `cNvPr-4/dsp/${index + 1}`, fingerprintSha256: `${index}`.repeat(64) },
    passthrough: [], compatibility: { status: 'preserveOnly', diagnostics: [diagnostic(id)] },
  }
}

function diagramDeck(children: NativeElement[] = ['Manager', 'Manager2', 'Assistant', 'Employee', 'Employee2'].map((text, index) => diagramShape(index, text, (index % 3) * 2_286_000, Math.floor(index / 3) * 1_371_600))): NativePptxDeck {
  const group: NativeElement = {
    kind: 'group', id: 'diagram-frame', provenance: 'parsed', name: 'Diagram 3',
    transform: { ...FRAME }, childTransform: { x: 0, y: 0, cx: FRAME.cx, cy: FRAME.cy }, children,
    source: { partName: SLIDE_PART, objectId: 'cNvPr-4', fingerprintSha256: 'f'.repeat(64) },
    passthrough: [{ token: 'diagram-frame-token', ownerPart: SLIDE_PART, fingerprintSha256: 'f'.repeat(64), disposition: 'preserve' }],
    compatibility: { status: 'preserveOnly', diagnostics: [diagnostic('diagram-frame')] },
  }
  return {
    contractVersion: 'pptx-native/v1', documentId: 'diagram-deck', origin: 'parsed', sourceRevision: 'rev-1', size: { cx: 9_144_000, cy: 6_858_000 }, assets: [],
    slides: [{ id: 'slide-1', provenance: 'parsed', elements: [group], source: { partName: SLIDE_PART, objectId: 'slide-1', fingerprintSha256: '1'.repeat(64) }, passthrough: [], compatibility: { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: 'slide.contains-preview', message: 'slide holds a read-only diagram preview', scope: { slideId: 'slide-1' } }] } }],
    compatibility: { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: 'deck.contains-preview', message: 'deck holds a read-only diagram preview' }] },
  }
}

/** Track the paint world matrix so we can read the absolute EMU origin of each painted path. */
function paintedPaths(commands: readonly PaintCommand[]): Array<{ fill?: string; stroke?: number; tx: number; ty: number }> {
  let matrix = [1, 0, 0, 1, 0, 0]
  const stack: number[][] = []
  const result: Array<{ fill?: string; stroke?: number; tx: number; ty: number }> = []
  for (const command of commands) {
    if (command.kind === 'save') stack.push([...matrix])
    if (command.kind === 'restore') matrix = stack.pop()!
    if (command.kind === 'transform') {
      const t = command.transform, [a, b, c, d, x, y] = matrix as [number, number, number, number, number, number]
      matrix = [(a * t.aPpm + c * t.bPpm) / 1e6, (b * t.aPpm + d * t.bPpm) / 1e6, (a * t.cPpm + c * t.dPpm) / 1e6, (b * t.cPpm + d * t.dPpm) / 1e6, a * t.txEmu + c * t.tyEmu + x, b * t.txEmu + d * t.tyEmu + y]
    }
    if (command.kind === 'path') result.push({ fill: command.fill, stroke: command.stroke?.widthEmu, tx: matrix[4]!, ty: matrix[5]! })
  }
  return result
}

it('paints diagram drawing fallback shapes as ordinary read-only paint at frame-relative EMU', async () => {
  const deck = diagramDeck(), before = JSON.stringify(deck)
  const validation = validateNativePptx(deck)
  expect(validation.ok, JSON.stringify(validation)).toBe(true)
  const tree = await compileNativePptxSlide(deck, 0, { textLayout, lineLayoutPolicy: 'max-run-natural-v1' })
  const surface = createRecordingPaintSurface(); paintSlideRenderTree(tree, surface)
  const commands = surface.finish()
  const paths = paintedPaths(commands).filter(path => path.fill === '2F6FED')
  expect(paths).toHaveLength(5)
  expect(paths.every(path => path.stroke === 12_700)).toBe(true)
  // dsp:sp offsets compose with the graphic frame origin through the identity child space.
  expect(paths[0]).toMatchObject({ tx: FRAME.x, ty: FRAME.y })
  expect(paths[1]).toMatchObject({ tx: FRAME.x + 2_286_000, ty: FRAME.y })
  expect(paths[3]).toMatchObject({ tx: FRAME.x, ty: FRAME.y + 1_371_600 })
  expect(tree.diagnostics.filter(item => item.severity === 'refusal')).toEqual([])
  expect(commands.filter(command => command.kind === 'glyphRun').length).toBeGreaterThanOrEqual(5)
  expect(commands.some(command => command.kind === 'placeholder')).toBe(false)
  expect(tree.diagnostics.filter(item => item.code === 'native.compatibility' && item.sourceCode === DIAGRAM_CODE)).toHaveLength(6)
  expect(tree.diagnostics.filter(item => item.code === 'render.preserveOnly').map(item => item.elementId)).toContain('diagram-frame')
  expect(JSON.stringify(deck)).toBe(before)
})

it('refuses diagram drawing shapes that claim editability or drop the declared preview warning', () => {
  const editable = diagramDeck()
  const group = editable.slides[0]!.elements[0]!
  if (group.kind !== 'group') throw new Error('group fixture')
  group.children[0]!.compatibility = { status: 'editable', diagnostics: [] }
  const issues = validateNativePptx(editable)
  expect(issues.ok).toBe(false)
  expect(JSON.stringify(issues)).toContain('native.geometryAuthority')

  const unlabeled = diagramDeck()
  const unlabeledGroup = unlabeled.slides[0]!.elements[0]!
  if (unlabeledGroup.kind !== 'group') throw new Error('group fixture')
  unlabeledGroup.children[1]!.compatibility = { status: 'preserveOnly', diagnostics: [] }
  expect(validateNativePptx(unlabeled).ok).toBe(false)
})

it('paints a refused diagram child as a placeholder instead of guessing its geometry', async () => {
  const refusedChild: NativeElement = {
    kind: 'shape', id: 'diagram-shape-refused', provenance: 'parsed', transform: { x: 0, y: 0, cx: 1_000_000, cy: 500_000 }, paragraphs: [],
    source: { partName: SLIDE_PART, objectId: 'cNvPr-4/dsp/9', fingerprintSha256: '9'.repeat(64) }, passthrough: [],
    compatibility: { status: 'refused', diagnostics: [{ severity: 'refusal', code: 'pptx.diagram-drawing-geometry-unavailable', message: 'geometry outside the evaluated profile', scope: { slideId: 'slide-1', elementId: 'diagram-shape-refused' } }] },
  }
  const deck = diagramDeck([diagramShape(0, 'Manager', 0, 0), refusedChild])
  const group = deck.slides[0]!.elements[0]!
  if (group.kind !== 'group') throw new Error('group fixture')
  group.compatibility = { status: 'refused', diagnostics: [{ severity: 'refusal', code: 'pptx.diagram-drawing-geometry-unavailable', message: 'child refused', scope: { slideId: 'slide-1', elementId: 'diagram-frame' } }] }
  deck.slides[0]!.compatibility.status = 'refused'
  deck.slides[0]!.compatibility.diagnostics.push({ severity: 'refusal', code: 'pptx.diagram-drawing-geometry-unavailable', message: 'slide holds a refused diagram', scope: { slideId: 'slide-1', elementId: 'diagram-frame' } })
  deck.compatibility = { status: 'refused', diagnostics: [{ severity: 'refusal', code: 'pptx.diagram-drawing-geometry-unavailable', message: 'deck holds a refused diagram' }] }
  const validation = validateNativePptx(deck)
  expect(validation.ok, JSON.stringify(validation)).toBe(true)
  const tree = await compileNativePptxSlide(deck, 0, { textLayout, lineLayoutPolicy: 'max-run-natural-v1' })
  const surface = createRecordingPaintSurface(); paintSlideRenderTree(tree, surface)
  const commands = surface.finish()
  expect(commands.some(command => command.kind === 'placeholder')).toBe(true)
  expect(commands.some(command => command.kind === 'path' && command.fill === '2F6FED')).toBe(false)
})
