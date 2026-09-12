import { createHash } from 'node:crypto'
import { readFileSync, readdirSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import type {
  NativeFontManifest,
  NativeFontResolver,
  NativeTextRefusal,
  NativeTextShaper,
  ResolvedFontFace,
  ShapedCluster,
  ShapedGlyph,
  ShapedSegment,
} from '@injoffice/font-metrics/layout'
import type { NativeElement, NativePptxDeck, NativeTextAlign, NativeTextBodyLayout } from '@injoffice/pptx-native'
import {
  RenderCompileError,
  compileNativePptxSlide,
  createRecordingPaintSurface,
  paintSlideRenderTree,
  paintSlideRenderTreeToCanvas2D,
  presetPath,
  stringifySlideRenderTree,
  type NativePptxTextLayout,
  type PaintCommand,
  type RenderNode,
  type RenderTextNode,
} from './index.js'

const root = resolve(import.meta.dirname, '../../..')
const parsedFull = JSON.parse(readFileSync(resolve(root, 'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'), 'utf8')) as NativePptxDeck
it('retains typed arrow source descriptors through compile and paint without mutation',async()=>{
 const deck=structuredClone(parsedFull),connector=deck.slides[0]!.elements.find(e=>e.kind==='connector')!
 if(connector.kind!=='connector')throw new Error('connector missing')
 connector.tailArrow=true;connector.tailEnd={type:'diamond',w:'lg',len:'sm'}
 const before=JSON.stringify(deck),tree=await compileNativePptxSlide(deck,0,{textLayout:textLayout()}),surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 expect(surface.finish()).toContainEqual(expect.objectContaining({kind:'path',tailArrow:true,tailEnd:{type:'diamond',w:'lg',len:'sm'}}));expect(JSON.stringify(deck)).toBe(before)
})
const digest = 'sha256:054edec1d0211f624fed0cbca9d4f9400b0e491c43742af2c5b0abebf0c990d8' as const

const manifest: NativeFontManifest = {
  version: 1,
  manifestId: 'render-fixture',
  revision: '1',
  faces: [{
    faceId: 'fixture.regular', family: 'Fixture Sans', aliases: ['Aptos'], weight: 400, style: 'normal', stretch: 100,
    source: { kind: 'bundled', resourceId: 'fixture-font', contentDigest: digest },
  }, {
    faceId: 'fixture.bold', family: 'Fixture Sans', aliases: ['Aptos'], weight: 700, style: 'normal', stretch: 100,
    source: { kind: 'bundled', resourceId: 'fixture-font-bold', contentDigest: digest },
  }],
  fallbackChains: [{ chainId: 'fixture.default', faceIds: ['fixture.regular', 'fixture.bold'] }],
}

function face(weight: number, matchedFamily = 'Aptos'): ResolvedFontFace {
  return {
    faceId: weight >= 700 ? 'fixture.bold' : 'fixture.regular', family: 'Fixture Sans', weight, style: 'normal', stretch: 100,
    sourceKind: 'bundled', resourceId: weight >= 700 ? 'fixture-font-bold' : 'fixture-font', contentDigest: digest,
    resolution: matchedFamily === 'Fixture Sans' ? 'exact' : 'substitute', matchedFamily, fallbackChainId: 'fixture.default',
  }
}

const resolver: NativeFontResolver = {
  providerId: 'fixture-resolver',
  providerRevision: '1',
  resolve({ run }) {
    return { status: 'resolved', face: face(run.font.weight, run.font.families[0]), attemptedFaceIds: ['fixture.regular'], decisions: [] }
  },
  load(resolved) {
    return {
      face: resolved,
      bytes: new Uint8Array([0, 1, 2, 3]),
      metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 },
    }
  },
}

function refusal(message: string): NativeTextRefusal {
  return { status: 'refused', attemptedFaceIds: ['fixture.regular'], decisions: [{ code: 'unsupported-script', message, recoverable: false }] }
}

function fixtureShaper(refuse: (text: string) => boolean = () => false): NativeTextShaper {
  return {
    providerId: 'fixture-shaper',
    providerRevision: '1',
    shape({ run, startUtf16, endUtf16, font }) {
      if (refuse(run.text)) return refusal('fixture refusal')
      const glyphs: ShapedGlyph[] = []
      const clusters: ShapedCluster[] = []
      let utf16 = startUtf16
      const logical: Array<{ character: string; startUtf16: number; endUtf16: number }> = []
      for (const character of run.text.slice(startUtf16, endUtf16)) {
        const end = utf16 + character.length
        logical.push({ character, startUtf16: utf16, endUtf16: end })
        utf16 = end
      }
      const paintOrder = run.direction === 'rtl' || run.direction === 'btt' ? [...logical].reverse() : logical
      for (const item of paintOrder) {
        const invisibleControl = /^[\u200b\u2060\ufeff]$/u.test(item.character)
        const glyphIndex = glyphs.length
        if (!invisibleControl) glyphs.push({ glyphId: item.character.codePointAt(0)!, clusterIndex: clusters.length, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 })
        clusters.push({ startUtf16: item.startUtf16, endUtf16: item.endUtf16, glyphStart: glyphIndex, glyphEnd: invisibleControl ? glyphIndex : glyphIndex + 1, advanceInlineMilliPoints: invisibleControl ? 0 : 1_000, whitespace: /^\s$/u.test(item.character) })
      }
      const ascent = run.fontSizeMilliPoints * 8 / 10
      const descent = -run.fontSizeMilliPoints * 2 / 10
      const lineGap = run.fontSizeMilliPoints * 2 / 10
      return {
        startUtf16, endUtf16, face: font.face, glyphs, clusters,
        metrics: { fontSizeMilliPoints: run.fontSizeMilliPoints, ascentMilliPoints: ascent, descentMilliPoints: descent, lineGapMilliPoints: lineGap, lineHeightMilliPoints: ascent - descent + lineGap },
        advanceInlineMilliPoints: clusters.reduce((sum, cluster) => sum + cluster.advanceInlineMilliPoints, 0),
        advanceBlockMilliPoints: 0,
      }
    },
  }
}

function textLayout(shaper = fixtureShaper(), resolveRun?: NativePptxTextLayout['resolveRun']): NativePptxTextLayout {
  return {
    manifest, resolver, shaper,
    defaults: { fontFamilies: ['Fixture Sans'], fontSizeHundredthPt: 1_000, script: 'Latn', language: 'en-US', direction: 'ltr', fallbackChainIds: ['fixture.default'] },
    resolveRun,
  }
}

it('never lays out source-frame autofit without opt-in and labels opted-in paint approximate',async()=>{
 const deck=structuredClone(parsedFull),element=deck.slides[0]!.elements.find(item=>item.kind==='text')!
 if(element.kind!=='text')throw new Error('text missing')
 const authored=nativeTextElement(element.id,'Saved frame text',nativeTextBody({autoFit:'shape-source-frame'}),element.transform)
 element.paragraphs=authored.paragraphs;element.textBody=authored.textBody
 element.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.autofit-source-frame-approximate',message:'Approximate saved frame'}]}
 deck.slides[0]!.elements=[element]
 const before=JSON.stringify(deck)
 const strict=await compileNativePptxSlide(deck,0,{textLayout:textLayout(),lineLayoutPolicy:'max-run-natural-v1'})
 expect(findNode(strict,'text',element.id).textBody).toMatchObject({status:'refused',fidelity:'nativeUnavailable'})
 const approximate=await compileNativePptxSlide(deck,0,{textLayout:textLayout(),lineLayoutPolicy:'max-run-natural-v1',sourceFrameAutoFitPreview:true})
 expect(findNode(approximate,'text',element.id).textBody).toMatchObject({status:'laidOut',fidelity:'approximateSourceFrame',autoFit:'shape-source-frame'})
 const paint=createRecordingPaintSurface();paintSlideRenderTree(approximate,paint)
 expect(paint.finish().some(command=>command.kind==='glyphRun')).toBe(true)
 expect(approximate.diagnostics.some(diagnostic=>diagnostic.code==='text.sourceFrameAutoFitApproximate')).toBe(true)
 expect(JSON.stringify(deck)).toBe(before)
 await expect(compileNativePptxSlide(deck,0,{textLayout:textLayout(),sourceFrameAutoFitPreview:'true' as never})).rejects.toMatchObject({path:'$.options.sourceFrameAutoFitPreview'})
})

it('passes authored language to actual native shaping and retains glyph paint', async () => {
  const base=fixtureShaper(),languages:string[]=[]
  const shaper:NativeTextShaper={...base,shape(request){languages.push(request.run.language);return base.shape(request)}}
  const element=nativeTextElement('authored-language','Istanbul',nativeTextBody(),{x:100,y:100,cx:1000000,cy:500000})
  element.paragraphs[0]!.runs[0]!.language='tr-TR'
  const tree=await compileNativePptxSlide(authoredDeck([element]),0,{textLayout:textLayout(shaper)})
  const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
  expect(languages.length).toBeGreaterThan(0);expect(languages.every(language=>language==='tr-TR')).toBe(true)
  expect(surface.finish().some(command=>command.kind==='glyphRun')).toBe(true)
  expect(surface.finish().some(command=>command.kind==='placeholder')).toBe(false)
})

function authoredDeck(elements: NativeElement[]): NativePptxDeck {
  return {
    contractVersion: 'pptx-native/v1', documentId: 'authored-render-deck', origin: 'authored',
    size: { cx: 2_000_000, cy: 1_500_000 }, assets: [],
    slides: [{ id: 'slide-authored', provenance: 'authored', elements, passthrough: [], compatibility: { status: 'editable', diagnostics: [] } }],
    compatibility: { status: 'editable', diagnostics: [] },
  }
}

function textElement(id: string, align: NativeTextAlign, y: number): Extract<NativeElement, { kind: 'text' }> {
  return {
    kind: 'text', id, provenance: 'authored', transform: { x: 100, y, cx: 1_000_000, cy: 200_000 },
    paragraphs: [{ align, level: 0, bullet: false, runs: [{ text: 'AB', bold: true, color: '112233' }, { text: 'CD', italic: true, color: '445566' }] }],
    passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
  }
}

function nativeTextBody(overrides: Partial<NativeTextBodyLayout> = {}): NativeTextBodyLayout {
  return {
    leftInsetEmu: 0, rightInsetEmu: 0, topInsetEmu: 0, bottomInsetEmu: 0,
    wrap: 'square', verticalAnchor: 'top', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow',
    ...overrides,
  }
}

it('applies native shape-frame quarter turns to glyph paint without reshaping horizontal text', async()=>{
 for(const quarterTurns of [1,2,3] as const){
  const element=nativeTextElement('rotated-text','AB',nativeTextBody(),{x:100,y:200,cx:500000,cy:500000})
  element.transform.quarterTurns=quarterTurns
  const tree=await compileNativePptxSlide(authoredDeck([element]),0,{textLayout:textLayout()})
  const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
  const commands=surface.finish()
  expect(commands.some(command=>command.kind==='glyphRun')).toBe(true)
  expect(commands.some(command=>command.kind==='placeholder')).toBe(false)
  expect(commands.some(command=>command.kind==='transform'&&(command.transform.aPpm!==1000000||command.transform.dPpm!==1000000))).toBe(true)
 }
})

it('retains source quarter-turn shape paths through native paint and checks rotated world bounds',async()=>{
 for(const quarterTurns of [1,2,3] as const){
  const shape:NativeElement={kind:'shape',id:'rotated',provenance:'authored',transform:{x:100,y:200,cx:400,cy:200,quarterTurns},preset:'triangle',fill:'123456',paragraphs:[],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}
  const tree=await compileNativePptxSlide(authoredDeck([shape]),0,{textLayout:textLayout()})
  const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
  expect(surface.finish().some(command=>command.kind==='path'&&command.fill==='123456')).toBe(true)
  expect(surface.finish().some(command=>command.kind==='placeholder')).toBe(false)
  shape.transform={x:0,y:900,cx:1000,cy:100,quarterTurns:1}
  const outside=authoredDeck([shape]);outside.size={cx:1000,cy:1000}
  await expect(compileNativePptxSlide(outside,0,{textLayout:textLayout(),maxCoordinateEmu:1000})).rejects.toMatchObject({code:'render.worldTransform'})
 }
})

it('composes scaled group coordinates with a rotated child and bounds every world corner',async()=>{
 const child:NativeElement={kind:'shape',id:'child-turn',provenance:'authored',transform:{x:100,y:200,cx:400,cy:200,quarterTurns:1},preset:'triangle',fill:'123456',paragraphs:[],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}
 const group:NativeElement={kind:'group',id:'scaled-parent',provenance:'authored',transform:{x:0,y:0,cx:2000,cy:3000},childTransform:{x:0,y:0,cx:1000,cy:1000},children:[child],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}
 const input=authoredDeck([group]);input.size={cx:4000,cy:4000}
 const tree=await compileNativePptxSlide(input,0,{textLayout:textLayout(),maxCoordinateEmu:4000})
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 let matrix=[1,0,0,1,0,0];const stack:number[][]=[];let painted:number[]|undefined
 for(const command of surface.finish()){
  if(command.kind==='save')stack.push([...matrix])
  if(command.kind==='restore')matrix=stack.pop()!
  if(command.kind==='transform'){
   const t=command.transform,[a,b,c,d,x,y]=matrix as [number,number,number,number,number,number]
   matrix=[(a*t.aPpm+c*t.bPpm)/1e6,(b*t.aPpm+d*t.bPpm)/1e6,(a*t.cPpm+c*t.dPpm)/1e6,(b*t.cPpm+d*t.dPpm)/1e6,a*t.txEmu+c*t.tyEmu+x,b*t.txEmu+d*t.tyEmu+y]
  }
  if(command.kind==='path'&&command.fill==='123456')painted=[...matrix]
 }
 expect(painted).toEqual([0,3,-2,0,800,300])
 child.transform={x:0,y:900,cx:1000,cy:100,quarterTurns:1}
 await expect(compileNativePptxSlide(input,0,{textLayout:textLayout(),maxCoordinateEmu:4000})).rejects.toMatchObject({code:'render.worldTransform'})
})

function nativeTextElement(id: string, text: string, body: NativeTextBodyLayout, transform = { x: 100, y: 0, cx: 500_000, cy: 500_000 }): Extract<NativeElement, { kind: 'text' }> {
  return {
    kind: 'text', id, provenance: 'authored', transform, textBody: body,
    paragraphs: [{ align: 'left', level: 0, bullet: false, runs: [{ text, fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }] }],
    passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
  }
}

function findNode<T extends RenderNode['kind']>(tree: Awaited<ReturnType<typeof compileNativePptxSlide>>, kind: T, id: string): Extract<RenderNode, { kind: T }> {
  const node = tree.nodes.find((candidate) => candidate.kind === kind && candidate.sourceElementId === id)
  if (!node) throw new Error(`missing ${kind} ${id}`)
  return node as Extract<RenderNode, { kind: T }>
}

describe('native PPTX RenderTree', () => {
  it('preserves strict refusal for marker layout and inherited paragraph margins', async () => {
    for (const override of [{ bullet: true, bulletCharacter: '▪' }, { marginLeftEmu: 300000 }, { indentEmu: -100000 }]) {
      const element = nativeTextElement('styled-paragraph', 'Hello', nativeTextBody())
      Object.assign(element.paragraphs[0]!, override)
      const tree = await compileNativePptxSlide(authoredDeck([element]), 0, { textLayout: textLayout() })
      expect(findNode(tree, 'text', element.id).textBody).toMatchObject({ fidelity: 'nativeUnavailable', status: 'refused', paragraphs: [] })
      expect(tree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.paragraphSemanticsUnavailable' })]))
    }
  })
  it('retains exact source crop through immutable image nodes and paint commands without rewriting assets', async () => {
    const deck = structuredClone(parsedFull)
    const picture = deck.slides[0]!.elements.find((item) => item.kind === 'picture')!
    if (picture.kind !== 'picture') throw new Error('fixture requires a picture')
    picture.crop = { left: 12500, top: 25000, right: 37500, bottom: 0 }
    const source = JSON.stringify(deck)
    const tree = await compileNativePptxSlide(deck, 0, { textLayout: textLayout() })
    const image = findNode(tree, 'image', picture.id)
    const asset = deck.assets.find((item) => item.id === picture.assetId)!
    expect(image).toMatchObject({ crop: picture.crop, sha256: asset.sha256, byteLength: asset.byteLength })
    expect(Object.isFrozen(image.crop)).toBe(true)
    const surface = createRecordingPaintSurface()
    paintSlideRenderTree(tree, surface)
    expect(surface.finish()).toContainEqual(expect.objectContaining({ kind: 'image', sourceElementId: picture.id, crop: picture.crop, rect: image.bounds }))
    expect(JSON.stringify(deck)).toBe(source)
    expect(stringifySlideRenderTree(await compileNativePptxSlide(deck, 0, { textLayout: textLayout() }))).toBe(stringifySlideRenderTree(tree))
  })
  it('does not paint an uncropped image when the extractor reports an unsupported source crop', async () => {
    const deck = structuredClone(parsedFull)
    const picture = deck.slides[0]!.elements.find((item) => item.kind === 'picture')!
    picture.compatibility = { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: 'pptx.picture-crop-unavailable', message: 'outset crop preserved' }] }
    const tree = await compileNativePptxSlide(deck, 0, { textLayout: textLayout() })
    expect(findNode(tree, 'placeholder', picture.id)).toMatchObject({ label: 'Unsupported picture crop preserved' })
    const surface = createRecordingPaintSurface()
    paintSlideRenderTree(tree, surface)
    expect(surface.finish().some((command) => command.kind === 'image' && command.sourceElementId === picture.id)).toBe(false)
  })

  it('compiles rich native content in stable z-order with explicit assets, clips, groups, tables, and chart preview', async () => {
    const tree = await compileNativePptxSlide(parsedFull, 'slide-a', { textLayout: textLayout() })
    expect(tree.nodes.map((node) => [node.zIndex, node.sourceElementId, node.kind])).toEqual([
      [0, 'el-title', 'text'], [1, 'el-shape', 'shape'], [2, 'el-line', 'connector'], [3, 'el-picture', 'image'],
      [4, 'el-table', 'table'], [5, 'el-chart', 'image'], [6, 'el-group', 'group'],
    ])
    const parsedGroup = findNode(tree, 'group', 'el-group')
    expect(parsedGroup.children[0]).toMatchObject({ sourceElementId: 'el-group-child', zIndex: 0, kind: 'shape' })
    expect(parsedGroup).not.toHaveProperty('clip')
    expect(parsedGroup).toMatchObject({
      transform: { aPpm: 1_000_000, dPpm: 1_000_000, txEmu: 7_000_000, tyEmu: 1_000_000 },
      bounds: { x: 0, y: 0, cx: 2_000_000, cy: 2_000_000 },
    })
    expect(findNode(tree, 'shape', 'el-shape').path).toEqual([{ kind: 'roundRect', rect: { x: 0, y: 0, cx: 1_500_000, cy: 800_000 }, radiusEmu: 100_000 }])
    expect(findNode(tree, 'connector', 'el-line')).toMatchObject({ headArrow: false, tailArrow: true })
    expect(findNode(tree, 'connector', 'el-line')).not.toHaveProperty('clip')
    expect(findNode(tree, 'image', 'el-picture')).toMatchObject({ role: 'picture', assetId: 'z-picture', resolutionSource: 'sourceDeck' })
    expect(findNode(tree, 'image', 'el-chart')).toMatchObject({ role: 'chartPreview', assetId: 'a-preview', resolutionSource: 'host' })
    expect(findNode(tree, 'table', 'el-table').cells).toHaveLength(4)
    expect(findNode(tree, 'table', 'el-table').cells[3]).toMatchObject({ bounds: { x: 1_000_000, y: 500_000, cx: 1_000_000, cy: 500_000 }, border: { color: '000000', widthEmu: 12_700 } })
    expect(tree.assets.map((asset) => asset.id)).toEqual(['a-preview', 'z-picture'])
    expect(tree.diagnostics.map((diagnostic) => diagnostic.code)).toEqual(expect.arrayContaining(['render.preserveOnly', 'asset.hostResolutionRequired', 'text.layoutMetadataUnavailable']))
    expect(Object.isFrozen(tree)).toBe(true)
    expect(Object.isFrozen(tree.nodes)).toBe(true)
    const canonical = stringifySlideRenderTree(tree)
    expect(stringifySlideRenderTree(await compileNativePptxSlide(parsedFull, 0, { textLayout: textLayout() }))).toBe(canonical)
    expect(createHash('sha256').update(canonical).digest('hex')).toBe('90501a2ed003730b6092d839e330fe7e9bab3591d8b8a941e4a08719454d0698')
  })

  it('compiles and paints exact table cells from renderer-neutral native text commands without cell or table clipping', async () => {
    const table: Extract<NativeElement, { kind: 'table' }> = {
      kind: 'table', id: 'native-table', provenance: 'authored',
      transform: { x: 100_000, y: 200_000, cx: 400_000, cy: 200_000 },
      table: {
        columnWidths: [150_000, 250_000], rowHeights: [200_000],
        rows: [[{
          text: 'AB', fill: 'AABBCC',
          paragraphs: [{ align: 'left', level: 0, bullet: false, runs: [
            { text: 'A', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000, color: '112233' },
            { text: 'B', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000, bold: true, color: '445566' },
          ] }],
          textBody: {
            leftInsetEmu: 10_000, rightInsetEmu: 20_000, topInsetEmu: 30_000, bottomInsetEmu: 40_000,
            wrap: 'square', verticalAnchor: 'top', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow',
          },
        }, {
          text: 'C', fill: 'DDEEFF',
          paragraphs: [{ align: 'right', level: 0, bullet: false, runs: [{ text: 'C', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000, italic: true }] }],
          textBody: {
            leftInsetEmu: 5_000, rightInsetEmu: 6_000, topInsetEmu: 7_000, bottomInsetEmu: 8_000,
            wrap: 'square', verticalAnchor: 'top', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow',
          },
        }]],
      },
      passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }
    const tree = await compileNativePptxSlide(authoredDeck([table]), 0, { textLayout: textLayout() })
    const node = findNode(tree, 'table', table.id)
    expect(node).not.toHaveProperty('clip')
    expect(node.cells.map((cell) => cell.bounds)).toEqual([
      { x: 0, y: 0, cx: 150_000, cy: 200_000 },
      { x: 150_000, y: 0, cx: 250_000, cy: 200_000 },
    ])
    expect(node.cells[0]).not.toHaveProperty('paragraph')
    expect(node.cells[0]!.textBody).toMatchObject({
      fidelity: 'native', status: 'laidOut', wrap: 'square', verticalAnchor: 'top',
      bounds: { x: 10_000, y: 30_000, cx: 120_000, cy: 130_000 },
    })
    expect(tree.diagnostics.map((diagnostic) => diagnostic.code)).not.toContain('text.layoutMetadataUnavailable')

    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(tree, recording)
    const commands = recording.finish()
    expect(commands.filter((command) => command.kind === 'clipRect')).toEqual([
      { kind: 'clipRect', rect: { x: 0, y: 0, cx: 2_000_000, cy: 1_500_000 } },
    ])
    expect(commands.filter((command) => command.kind === 'path')).toEqual([
      expect.objectContaining({ kind: 'path', sourceElementId: 'native-table', fill: 'AABBCC' }),
      expect.objectContaining({ kind: 'path', sourceElementId: 'native-table', fill: 'DDEEFF' }),
    ])
    expect(commands.filter((command) => command.kind === 'glyphRun').map((command) => command.kind === 'glyphRun' ? command.run.text : '')).toEqual(['A', 'B', 'C'])
    const secondFill = commands.findIndex((command) => command.kind === 'path' && command.fill === 'DDEEFF')
    const firstGlyph = commands.findIndex((command) => command.kind === 'glyphRun' && command.run.text === 'A')
    expect(secondFill).toBeGreaterThan(-1)
    expect(firstGlyph).toBeGreaterThan(secondFill)

    const grouped: NativeElement = {
      kind: 'group', id: 'table-group', provenance: 'authored',
      transform: { x: 1_000_000, y: 2_000_000, cx: 800_000, cy: 400_000 },
      childTransform: { x: 0, y: 0, cx: 400_000, cy: 200_000 }, children: [table],
      passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }
    const groupedTree = await compileNativePptxSlide(authoredDeck([grouped]), 0, { textLayout: textLayout() })
    const groupedNode = groupedTree.nodes[0]
    expect(groupedNode).toMatchObject({ kind: 'group', transform: { aPpm: 2_000_000, dPpm: 2_000_000, txEmu: 1_000_000, tyEmu: 2_000_000 } })
    if (!groupedNode || groupedNode.kind !== 'group') throw new Error('missing grouped table')
    expect(groupedNode.children[0]).toMatchObject({ kind: 'table', sourceElementId: 'native-table', transform: { txEmu: 100_000, tyEmu: 200_000 } })
    const groupedRecording = createRecordingPaintSurface()
    paintSlideRenderTree(groupedTree, groupedRecording)
    expect(groupedRecording.finish().filter((command) => command.kind === 'transform').slice(0, 4)).toEqual([
      { kind: 'transform', transform: { aPpm: 2_000_000, bPpm: 0, cPpm: 0, dPpm: 2_000_000, txEmu: 1_000_000, tyEmu: 2_000_000 } },
      { kind: 'transform', transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: 100_000, tyEmu: 200_000 } },
      { kind: 'transform', transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: 0, tyEmu: 0 } },
      { kind: 'transform', transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: 150_000, tyEmu: 0 } },
    ])
    expect(readFileSync(resolve(root, 'packages/pptx-render/src/paint.ts'), 'utf8')).not.toMatch(/\b(?:document|window|HTMLElement|innerHTML|DOMParser)\b/)
  })

  it('keeps nested native group projection renderer-neutral and composes source order without group clipping', async () => {
    const nested: NativeElement = {
      kind: 'group', id: 'outer-native-group', provenance: 'authored',
      transform: { x: 1_000_000, y: 2_000_000, cx: 6_000_000, cy: 8_000_000 },
      childTransform: { x: 100, y: 200, cx: 300, cy: 400 },
      children: [{
        kind: 'group', id: 'inner-native-group', provenance: 'authored',
        transform: { x: 150, y: 250, cx: 100, cy: 100 },
        childTransform: { x: 10, y: 20, cx: 50, cy: 50 },
        children: [{
          kind: 'shape', id: 'nested-native-shape', provenance: 'authored',
          transform: { x: 20, y: 30, cx: 10, cy: 5 }, preset: 'rect', paragraphs: [],
          stroke: { color: '112233', widthEmu: 2 }, passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
        }, {
          kind: 'text', id: 'nested-native-text', provenance: 'authored',
          transform: { x: 22, y: 32, cx: 10, cy: 5 }, paragraphs: [{ runs: [{ text: 'A' }] }],
          passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
        }, {
          kind: 'picture', id: 'nested-native-picture', provenance: 'authored', assetId: 'nested-picture-asset',
          transform: { x: 24, y: 34, cx: 10, cy: 5 }, passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
        }],
        passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
      }],
      passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }
    const deck = authoredDeck([nested])
    deck.assets.push({
      id: 'nested-picture-asset', provenance: 'authored', contentType: 'image/png',
      sha256: '0'.repeat(64), byteLength: 0, passthrough: [],
    })
    const tree = await compileNativePptxSlide(deck, 0, { textLayout: textLayout() })
    const outer = findNode(tree, 'group', nested.id)
    const inner = outer.children[0]!
    const children = inner.kind === 'group' ? inner.children : []
    expect(outer).not.toHaveProperty('clip')
    expect(outer).toMatchObject({ transform: { aPpm: 20_000_000_000, dPpm: 20_000_000_000, txEmu: -1_000_000, tyEmu: -2_000_000 } })
    expect(inner).not.toHaveProperty('clip')
    expect(inner).toMatchObject({ kind: 'group', transform: { aPpm: 2_000_000, dPpm: 2_000_000, txEmu: 130, tyEmu: 210 } })
    expect(children.map((child) => [child.sourceElementId, child.kind])).toEqual([
      ['nested-native-shape', 'shape'], ['nested-native-text', 'text'], ['nested-native-picture', 'image'],
    ])

    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(tree, recording)
    const ppm = 1_000_000n
    type World = { a: bigint; d: bigint; tx: bigint; ty: bigint }
    let world: World = { a: ppm, d: ppm, tx: 0n, ty: 0n }
    const stack: World[] = []
    const painted: Array<{ kind: string; id: string; world: World; strokeWidth?: number }> = []
    for (const command of recording.finish()) {
      if (command.kind === 'save') stack.push(world)
      else if (command.kind === 'restore') world = stack.pop()!
      else if (command.kind === 'transform') {
        const nextA = BigInt(command.transform.aPpm)
        const nextD = BigInt(command.transform.dPpm)
        world = {
          a: world.a * nextA / ppm,
          d: world.d * nextD / ppm,
          tx: world.tx + world.a * BigInt(command.transform.txEmu) / ppm,
          ty: world.ty + world.d * BigInt(command.transform.tyEmu) / ppm,
        }
      } else if (command.kind === 'path' || command.kind === 'glyphRun' || command.kind === 'image') {
        painted.push({ kind: command.kind, id: command.sourceElementId, world: { ...world }, strokeWidth: command.kind === 'path' ? command.stroke?.widthEmu : undefined })
      }
    }
    expect(painted).toEqual(expect.arrayContaining([
      { kind: 'path', id: 'nested-native-shape', world: { a: 40_000_000_000n, d: 40_000_000_000n, tx: 2_400_000n, ty: 3_400_000n }, strokeWidth: 2 },
      { kind: 'glyphRun', id: 'nested-native-text', world: { a: 40_000_000_000n, d: 40_000_000_000n, tx: 2_480_000n, ty: 3_480_000n }, strokeWidth: undefined },
      { kind: 'image', id: 'nested-native-picture', world: { a: 40_000_000_000n, d: 40_000_000_000n, tx: 2_560_000n, ty: 3_560_000n }, strokeWidth: undefined },
    ]))
  })

  it('retains exact rational cumulative group transforms and refuses only out-of-budget world space', async () => {
    const leaf: NativeElement = {
      kind: 'shape', id: 'world-leaf', provenance: 'authored', preset: 'rect', paragraphs: [],
      transform: { x: 0, y: 0, cx: 1, cy: 1 }, passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }
    const wrap = (child: NativeElement, id: string, extent: number, childExtent: number): NativeElement => ({
      kind: 'group', id, provenance: 'authored', transform: { x: 0, y: 0, cx: extent, cy: extent },
      childTransform: { x: 0, y: 0, cx: childExtent, cy: childExtent }, children: [child],
      passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    })

    const fractional = wrap(wrap(leaf, 'fractional-inner', 1_000_001, 1_000_000), 'fractional-outer', 1_000_001, 1_000_000)
    await expect(compileNativePptxSlide(authoredDeck([fractional]), 0, { textLayout: textLayout() })).resolves.toMatchObject({ nodes: [{ kind: 'group' }] })

    const halfEmuLeaf: NativeElement = {
      ...leaf, id: 'half-emu-leaf', transform: { x: 1, y: 1, cx: 1, cy: 1 },
    }
    const halfEmu = wrap(halfEmuLeaf, 'half-emu-group', 3, 2)
    await expect(compileNativePptxSlide(authoredDeck([halfEmu]), 0, { textLayout: textLayout() })).resolves.toMatchObject({ nodes: [{ kind: 'group' }] })

    const largeSafeCoefficient = wrap(leaf, 'large-safe-coefficient', 1_000_000_000, 1)
    await expect(compileNativePptxSlide(authoredDeck([largeSafeCoefficient]), 0, { textLayout: textLayout() })).resolves.toMatchObject({ nodes: [{ kind: 'group' }] })

    let excessive: NativeElement = leaf
    for (let depth = 0; depth < 5; depth++) excessive = wrap(excessive, `bounded-world-${depth}`, 1_000, 1)
    await expect(compileNativePptxSlide(authoredDeck([excessive]), 0, { textLayout: textLayout() })).rejects.toMatchObject({ code: 'render.worldTransform' })
  })

  it('places mixed rich runs using exact shaped advances for all LTR and RTL alignments', async () => {
    const elements = [textElement('ltr-left', 'left', 0), textElement('ltr-center', 'center', 250_000), textElement('ltr-right', 'right', 500_000), textElement('rtl-left', 'left', 750_000), textElement('rtl-center', 'center', 1_000_000), textElement('rtl-right', 'right', 1_250_000)]
    const tree = await compileNativePptxSlide(authoredDeck(elements), 0, {
      textLayout: textLayout(fixtureShaper(), ({ elementId }) => elementId.startsWith('rtl-') ? { direction: 'rtl', script: 'Arab', language: 'ar' } : {}),
    })
    const positions = (id: string) => (findNode(tree, 'text', id) as RenderTextNode).textBody.paragraphs[0]!.runs.map((run) => run.x)
    expect(positions('ltr-left')).toEqual([0, 25_400])
    expect(positions('ltr-center')).toEqual([474_600, 500_000])
    expect(positions('ltr-right')).toEqual([949_200, 974_600])
    expect(positions('rtl-left')).toEqual([25_400, 0])
    expect(positions('rtl-center')).toEqual([500_000, 474_600])
    expect(positions('rtl-right')).toEqual([974_600, 949_200])
    expect(findNode(tree, 'text', 'rtl-right').textBody.paragraphs[0]!.direction).toBe('rtl')
    expect(findNode(tree, 'text', 'rtl-right').textBody.paragraphs[0]!.runs[0]!.clusters.map((cluster) => [cluster.startUtf16, cluster.endUtf16])).toEqual([[1, 2], [0, 1]])

    const mixed = await compileNativePptxSlide(authoredDeck([textElement('mixed-direction', 'left', 0)]), 0, {
      textLayout: textLayout(fixtureShaper(), ({ runIndex }) => runIndex === 1 ? { direction: 'rtl', script: 'Arab', language: 'ar' } : {}),
    })
    expect(mixed.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'text.bidiUnavailable', elementId: 'mixed-direction' }),
    ]))
  })

  it('projects exact native insets and top anchoring, but visibly refuses unqualified center/bottom anchoring', async () => {
    const elements: NativeElement[] = (['top', 'center', 'bottom'] as const).map((verticalAnchor, index) => nativeTextElement(
      `anchor-${verticalAnchor}`,
      'AB',
      nativeTextBody({ leftInsetEmu: 10_000, rightInsetEmu: 20_000, topInsetEmu: 30_000, bottomInsetEmu: 40_000, wrap: 'none', verticalAnchor }),
      { x: 100, y: index * 600_000, cx: 500_000, cy: 500_000 },
    ))
    elements.push({
      kind: 'shape', id: 'text-bearing-shape', provenance: 'authored',
      transform: { x: 600_000, y: 0, cx: 500_000, cy: 500_000 }, preset: 'rect',
      textBody: nativeTextBody({ leftInsetEmu: 10_000, rightInsetEmu: 20_000, topInsetEmu: 30_000, bottomInsetEmu: 40_000, wrap: 'none' }),
      paragraphs: [{ align: 'left', level: 0, bullet: false, runs: [{ text: 'shape' }] }], passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    })
    const tree = await compileNativePptxSlide(authoredDeck(elements), 0, { textLayout: textLayout() })
    for (const verticalAnchor of ['top', 'center', 'bottom'] as const) {
      const node = findNode(tree, 'text', `anchor-${verticalAnchor}`)
      expect(node.clip).toBeUndefined()
      expect(node.textBody).toMatchObject({
        fidelity: verticalAnchor === 'top' ? 'native' : 'nativeUnavailable', status: verticalAnchor === 'top' ? 'laidOut' : 'refused', wrap: 'none', verticalAnchor,
        autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow',
        bounds: { x: 10_000, y: 30_000, cx: 470_000, cy: 430_000 },
      })
      if (verticalAnchor === 'top') expect(node.textBody.paragraphs[0]).toMatchObject({ y: 30_000, heightEmu: 152_400 })
      else expect(node.textBody).toMatchObject({ paragraphs: [], refusalLabel: 'Exact text layout unavailable' })
    }
    expect(tree.diagnostics.filter((item) => item.code === 'text.verticalAnchorUnavailable').map((item) => item.elementId).sort()).toEqual(['anchor-bottom', 'anchor-center'])
    const shape = findNode(tree, 'shape', 'text-bearing-shape')
    expect(shape.clip).toBeUndefined()
    expect(shape.textBody?.bounds).toEqual({ x: 10_000, y: 30_000, cx: 470_000, cy: 430_000 })
  })

  it('opts into measured mixed-size lines and integer anchor placement without changing strict defaults', async () => {
    const elements = (['top','center','bottom'] as const).map(anchor => {
      const element = nativeTextElement(`mixed-${anchor}`, 'AB', nativeTextBody({wrap:'none',verticalAnchor:anchor}), {x:0,y:0,cx:500000,cy:500001})
      return {...element, paragraphs:[{align:'left' as const,level:0,bullet:false,runs:[{text:'A',fontSizeHundredthPt:1000},{text:'B',fontSizeHundredthPt:2000}]}]}
    })
    const deck=authoredDeck(elements), before=JSON.stringify(deck)
    const strict=await compileNativePptxSlide(deck,0,{textLayout:textLayout()})
    for(const element of elements) expect(findNode(strict,'text',element.id).textBody.status).toBe('refused')
    const tree=await compileNativePptxSlide(deck,0,{textLayout:textLayout(),lineLayoutPolicy:'max-run-natural-v1'})
    for(const [index,element] of elements.entries()) {
      const body=findNode(tree,'text',element.id).textBody
      expect(body).toMatchObject({fidelity:'deterministicNative',lineLayoutPolicy:'max-run-natural-v1',status:'laidOut'})
      const line=body.paragraphs[0]!
      expect(line.heightEmu).toBe(304800)
      expect(line.y).toBe([0,97600,195201][index])
      expect(line.runs.map(run=>run.baselineY)).toEqual([line.y+203200,line.y+203200])
    }
    expect(JSON.stringify(deck)).toBe(before)
    expect(tree.diagnostics.filter(d=>d.code==='text.deterministicLayout')).toHaveLength(3)
    await expect(compileNativePptxSlide(deck,0,{textLayout:textLayout(),lineLayoutPolicy:'unknown' as never})).rejects.toMatchObject({path:'$.options.lineLayoutPolicy'})
  })

  it('anchors the complete multiline block with signed overflow under the named measured policy', async () => {
    const element=nativeTextElement('overflow-center','A',nativeTextBody({wrap:'none',verticalAnchor:'center'}),{x:0,y:0,cx:500000,cy:100001})
    element.paragraphs=[...element.paragraphs,...element.paragraphs]
    const tree=await compileNativePptxSlide(authoredDeck([element]),0,{textLayout:textLayout(),lineLayoutPolicy:'max-run-natural-v1'})
    const lines=findNode(tree,'text',element.id).textBody.paragraphs
    expect(lines.map(line=>line.y)).toEqual([-102400,50000])
    expect(lines.map(line=>line.runs[0]!.baselineY)).toEqual([-800,151600])
  })

  it('shapes one authored marker at its hanging indent and preserves content-run source ranges across wrapping', async () => {
    const element=nativeTextElement('bullet','AB CD',nativeTextBody(),{x:0,y:0,cx:38100,cy:500000})
    element.paragraphs=[{...element.paragraphs[0]!,level:2,bullet:true,bulletCharacter:'▪',marginLeftEmu:12700,indentEmu:-12700}]
    const deck=authoredDeck([element]),before=JSON.stringify(deck)
    const strict=await compileNativePptxSlide(deck,0,{textLayout:textLayout()})
    expect(findNode(strict,'text',element.id).textBody.status).toBe('refused')
    const tree=await compileNativePptxSlide(deck,0,{textLayout:textLayout(),lineLayoutPolicy:'max-run-natural-v1'})
    const lines=findNode(tree,'text',element.id).textBody.paragraphs
    expect(lines).toHaveLength(2)
    expect(lines.map(line=>line.x)).toEqual([12700,12700])
    expect(lines[0]!.marker).toMatchObject({sourceRole:'paragraphBullet',text:'▪',x:0,baselineY:101600,startUtf16:0,endUtf16:1})
    expect(lines[1]!.marker).toBeUndefined()
    expect(lines.flatMap(line=>line.runs.map(run=>[run.text,run.startUtf16,run.endUtf16]))).toEqual([['AB',0,2],['CD',3,5]])
    const paint=createRecordingPaintSurface();paintSlideRenderTree(tree,paint)
    expect(paint.finish().filter(command=>command.kind==='glyphRun').map(command=>command.kind==='glyphRun'?command.run.text:'')).toEqual(['▪','AB','CD'])
    expect(JSON.stringify(deck)).toBe(before)
  })

  it('uses different first/continuation widths for a positive non-list indent and refuses ambiguous bullet geometry', async () => {
    const base=nativeTextElement('indent','AB CD',nativeTextBody(),{x:0,y:0,cx:50800,cy:500000})
    base.paragraphs=[{...base.paragraphs[0]!,marginLeftEmu:12700,indentEmu:12700}]
    const tree=await compileNativePptxSlide(authoredDeck([base]),0,{textLayout:textLayout(),lineLayoutPolicy:'max-run-natural-v1'})
    expect(findNode(tree,'text',base.id).textBody.paragraphs.map(p=>p.x)).toEqual([25400,12700])
    for(const override of [{bullet:true,bulletCharacter:'▪',indentEmu:0},{bullet:true,bulletCharacter:'▪',indentEmu:-12700,align:'center' as const},{level:2,marginLeftEmu:undefined}]){
      const element={...base,paragraphs:[{...base.paragraphs[0]!,...override}]}
      if(element.paragraphs[0]!.marginLeftEmu===undefined)delete element.paragraphs[0]!.marginLeftEmu
      const refused=await compileNativePptxSlide(authoredDeck([element]),0,{textLayout:textLayout(),lineLayoutPolicy:'max-run-natural-v1'})
      expect(findNode(refused,'text',base.id).textBody.status).toBe('refused')
    }
  })

  it('wraps only at modeled shaped-cluster boundaries and leaves native no-wrap text on one overflowing line', async () => {
    const wrapped = nativeTextElement('cluster-wrap', 'AB CD', nativeTextBody(), { x: 0, y: 0, cx: 38_100, cy: 500_000 })
    const noWrap = nativeTextElement('cluster-no-wrap', 'A💡 B', nativeTextBody({ wrap: 'none' }), { x: 0, y: 600_000, cx: 25_400, cy: 500_000 })
    const tree = await compileNativePptxSlide(authoredDeck([wrapped, noWrap]), 0, { textLayout: textLayout() })
    const lines = findNode(tree, 'text', wrapped.id).textBody.paragraphs
    expect(lines).toHaveLength(2)
    expect(lines.map((line) => line.lineIndex)).toEqual([0, 1])
    expect(lines.map((line) => line.runs.map((run) => [run.text, run.startUtf16, run.endUtf16]))).toEqual([
      [['AB', 0, 2]],
      [['CD', 3, 5]],
    ])
    expect(lines[1]!.runs[0]!.clusters.map((cluster) => [cluster.startUtf16, cluster.endUtf16])).toEqual([[0, 1], [1, 2]])
    const whole = findNode(tree, 'text', noWrap.id).textBody
    expect(whole.paragraphs).toHaveLength(1)
    expect(whole.paragraphs[0]!.runs[0]).toMatchObject({ text: 'A💡 B', startUtf16: 0, endUtf16: 5 })
    expect(tree.diagnostics.find((diagnostic) => diagnostic.elementId === noWrap.id && diagnostic.code === 'text.overflow')).toBeUndefined()
  })

  it('consumes only an exact soft U+0020 separator and conservatively preserves or refuses ambiguous spacing', async () => {
    const exact = nativeTextElement('space-exact-fit', 'A B', nativeTextBody(), { x: 0, y: 0, cx: 12_700, cy: 500_000 })
    const centered = nativeTextElement('space-center', 'A B', nativeTextBody(), { x: 0, y: 0, cx: 20_000, cy: 500_000 })
    centered.paragraphs[0]!.align = 'center'
    const right = nativeTextElement('space-right', 'A B', nativeTextBody(), { x: 0, y: 0, cx: 20_000, cy: 500_000 })
    right.paragraphs[0]!.align = 'right'
    const fitting = [
      nativeTextElement('preserved-leading-fit', ' A', nativeTextBody(), { x: 0, y: 0, cx: 25_400, cy: 500_000 }),
      nativeTextElement('preserved-trailing-fit', 'A ', nativeTextBody(), { x: 0, y: 0, cx: 25_400, cy: 500_000 }),
      nativeTextElement('preserved-consecutive-fit', 'A  B', nativeTextBody(), { x: 0, y: 0, cx: 50_800, cy: 500_000 }),
      nativeTextElement('ideographic-space-fit', 'A\u3000B', nativeTextBody(), { x: 0, y: 0, cx: 38_100, cy: 500_000 }),
    ]
    const ambiguous = [
      nativeTextElement('preserved-leading-wrap', ' A', nativeTextBody(), { x: 0, y: 0, cx: 20_000, cy: 500_000 }),
      nativeTextElement('preserved-trailing-wrap', 'A ', nativeTextBody(), { x: 0, y: 0, cx: 12_700, cy: 500_000 }),
      nativeTextElement('preserved-consecutive-wrap', 'A  B', nativeTextBody(), { x: 0, y: 0, cx: 20_000, cy: 500_000 }),
      nativeTextElement('ideographic-space-wrap', 'A\u3000B', nativeTextBody(), { x: 0, y: 0, cx: 20_000, cy: 500_000 }),
    ]
    const tree = await compileNativePptxSlide(authoredDeck([exact, centered, right, ...fitting, ...ambiguous]), 0, { textLayout: textLayout() })

    for (const id of [exact.id, centered.id, right.id]) {
      const body = findNode(tree, 'text', id).textBody
      expect(body.paragraphs.map((line) => line.runs.map((run) => [run.text, run.startUtf16, run.endUtf16]))).toEqual([
        [['A', 0, 1]],
        [['B', 2, 3]],
      ])
      expect(body.paragraphs[0]!.consumedSoftSeparators).toEqual([{ sourceElementId: id, paragraphIndex: 0, runIndex: 0, startUtf16: 1, endUtf16: 2 }])
      expect(body.paragraphs.flatMap((line) => line.runs).flatMap((run) => run.glyphs).map((glyph) => glyph.glyphId)).not.toContain(0x20)
    }
    expect(findNode(tree, 'text', centered.id).textBody.paragraphs.map((line) => [line.x, line.widthEmu])).toEqual([[3_650, 12_700], [3_650, 12_700]])
    expect(findNode(tree, 'text', right.id).textBody.paragraphs.map((line) => [line.x, line.widthEmu])).toEqual([[7_300, 12_700], [7_300, 12_700]])

    for (const element of fitting) {
      const body = findNode(tree, 'text', element.id).textBody
      expect(body).toMatchObject({ status: 'laidOut' })
      expect(body.paragraphs).toHaveLength(1)
      expect(body.paragraphs[0]!.runs[0]!.text).toBe(element.paragraphs[0]!.runs[0]!.text)
    }
    for (const element of ambiguous) {
      expect(findNode(tree, 'text', element.id).textBody).toMatchObject({ fidelity: 'nativeUnavailable', status: 'refused', paragraphs: [] })
      expect(tree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.wrapUnavailable', elementId: element.id })]))
    }
  })

  it('rebases both glyph pen axes monotonically when a wrapped fragment skips a consumed separator', async () => {
    const base = fixtureShaper()
    const twoAxisPen: NativeTextShaper = {
      ...base,
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result) return result
        return {
          ...result,
          glyphs: result.glyphs.map((glyph) => ({ ...glyph, advanceYMilliPoints: 100 })),
        }
      },
    }
    const element = nativeTextElement('two-axis-fragment-pen', 'A B', nativeTextBody(), { x: 0, y: 0, cx: 12_700, cy: 500_000 })
    const tree = await compileNativePptxSlide(authoredDeck([element]), 0, { textLayout: textLayout(twoAxisPen) })
    const runs = findNode(tree, 'text', element.id).textBody.paragraphs.flatMap((line) => line.runs)
    expect(runs.map((run) => run.text)).toEqual(['A', 'B'])
    expect(runs.map((run) => run.glyphs.map((glyph) => [glyph.xEmu, glyph.yEmu]))).toEqual([[[0, 0]], [[0, 0]]])
  })

  it('uses shared Unicode cluster text rules for glue, WJ, ZWSP, hyphen/slash, and CJK punctuation', async () => {
    const base = fixtureShaper()
    const maliciousWhitespaceHints: NativeTextShaper = {
      ...base,
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result) return result
        return { ...result, clusters: result.clusters.map((cluster) => ({ ...cluster, whitespace: true })) }
      },
    }
    const cases = [
      nativeTextElement('nbsp', 'A\u00a0B', nativeTextBody(), { x: 0, y: 0, cx: 25_400, cy: 200_000 }),
      nativeTextElement('word-joiner', 'A\u2060B', nativeTextBody(), { x: 0, y: 220_000, cx: 25_400, cy: 200_000 }),
      nativeTextElement('zero-width-space', 'A\u200bB', nativeTextBody(), { x: 0, y: 440_000, cx: 12_700, cy: 300_000 }),
      nativeTextElement('hyphen', 'A-B', nativeTextBody(), { x: 0, y: 760_000, cx: 25_400, cy: 300_000 }),
      nativeTextElement('slash', 'A/B', nativeTextBody(), { x: 0, y: 1_080_000, cx: 25_400, cy: 300_000 }),
      nativeTextElement('cjk-close', '\u6f22\u3001\u5b57\u8a9e', nativeTextBody(), { x: 600_000, y: 0, cx: 25_400, cy: 300_000 }),
      nativeTextElement('cjk-open', '\u6f22\u300c\u5b57\u8a9e', nativeTextBody(), { x: 600_000, y: 320_000, cx: 25_400, cy: 500_000 }),
      nativeTextElement('unknown-break-class', 'A💡 B', nativeTextBody(), { x: 600_000, y: 840_000, cx: 25_400, cy: 300_000 }),
    ]
    const tree = await compileNativePptxSlide(authoredDeck(cases), 0, { textLayout: textLayout(maliciousWhitespaceHints) })

    for (const id of ['nbsp', 'unknown-break-class']) {
      expect(findNode(tree, 'text', id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    }
    expect(findNode(tree, 'text', 'word-joiner').textBody).toMatchObject({ status: 'laidOut' })
    // The fixture provider deliberately labels NBSP as whitespace; actual
    // cluster text wins, so that hostile/mistaken hint cannot create a break.
    expect(tree.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'text.wrapUnavailable', elementId: 'nbsp' }),
      expect.objectContaining({ code: 'text.wrapUnavailable', elementId: 'unknown-break-class' }),
    ]))
    expect(findNode(tree, 'text', 'zero-width-space').textBody.paragraphs.map((line) => line.runs.map((run) => run.text))).toEqual([['A\u200b'], ['B']])
    expect(findNode(tree, 'text', 'hyphen').textBody.paragraphs.map((line) => line.runs.map((run) => run.text))).toEqual([['A-'], ['B']])
    expect(findNode(tree, 'text', 'slash').textBody.paragraphs.map((line) => line.runs.map((run) => run.text))).toEqual([['A/'], ['B']])
    expect(findNode(tree, 'text', 'cjk-close').textBody.paragraphs.map((line) => line.runs.map((run) => run.text))).toEqual([['\u6f22\u3001'], ['\u5b57\u8a9e']])
    expect(findNode(tree, 'text', 'cjk-open').textBody.paragraphs.map((line) => line.runs.map((run) => run.text))).toEqual([['\u6f22'], ['\u300c\u5b57'], ['\u8a9e']])
  })

  it('visibly refuses mixed line metrics until Office-qualified leading aggregation is available', async () => {
    const base = fixtureShaper()
    const mixedMetrics: NativeTextShaper = {
      ...base,
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result) return result
        const metrics = request.run.text === 'A'
          ? { fontSizeMilliPoints: request.run.fontSizeMilliPoints, ascentMilliPoints: 10_000, descentMilliPoints: -1_000, lineGapMilliPoints: 0, lineHeightMilliPoints: 11_000 }
          : request.run.text === 'B'
            ? { fontSizeMilliPoints: request.run.fontSizeMilliPoints, ascentMilliPoints: 5_000, descentMilliPoints: -4_000, lineGapMilliPoints: 3_000, lineHeightMilliPoints: 12_000 }
            : result.metrics
        return { ...result, metrics }
      },
    }
    const element = nativeTextElement('mixed-line-metrics', 'unused', nativeTextBody({ wrap: 'none' }))
    element.paragraphs = [
      { align: 'left', level: 0, bullet: false, runs: [{ text: 'A', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }, { text: 'B', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }] },
      { align: 'left', level: 0, bullet: false, runs: [{ text: 'C', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }] },
    ]
    const tree = await compileNativePptxSlide(authoredDeck([element]), 0, { textLayout: textLayout(mixedMetrics) })
    expect(findNode(tree, 'text', element.id).textBody).toMatchObject({ fidelity: 'nativeUnavailable', status: 'refused', paragraphs: [], refusalLabel: 'Exact text layout unavailable' })
    expect(tree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.refused', severity: 'refusal', elementId: element.id })]))
  })

  it('retains exact descending RTL clusters for no-wrap but refuses RTL square wrap and mixed native directions', async () => {
    const rtlSquare = nativeTextElement('rtl-square', 'AB', nativeTextBody())
    const rtlNoWrap = nativeTextElement('rtl-no-wrap', 'AB', nativeTextBody({ wrap: 'none' }))
    const mixed = { ...textElement('native-mixed', 'left', 0), textBody: nativeTextBody({ wrap: 'none' }) }
    const tree = await compileNativePptxSlide(authoredDeck([rtlSquare, rtlNoWrap, mixed]), 0, {
      textLayout: textLayout(fixtureShaper(), ({ elementId, runIndex }) => {
        if (elementId === rtlSquare.id || elementId === rtlNoWrap.id || (elementId === mixed.id && runIndex === 1)) return { direction: 'rtl', script: 'Arab', language: 'ar' }
        return {}
      }),
    })
    expect(findNode(tree, 'text', rtlSquare.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    expect(findNode(tree, 'text', mixed.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    const exactRtl = findNode(tree, 'text', rtlNoWrap.id).textBody
    expect(exactRtl.status).toBe('laidOut')
    expect(exactRtl.paragraphs[0]).toMatchObject({ direction: 'rtl', lineIndex: 0 })
    expect(exactRtl.paragraphs[0]!.runs[0]!.clusters.map((cluster) => [cluster.startUtf16, cluster.endUtf16])).toEqual([[1, 2], [0, 1]])
    expect(tree.diagnostics.filter((diagnostic) => diagnostic.code === 'text.wrapUnavailable').map((diagnostic) => diagnostic.elementId).sort()).toEqual(['native-mixed', 'rtl-square'])
  })

  it('refuses an overfull unbreakable shaped cluster visibly instead of splitting or approximating it', async () => {
    const base = fixtureShaper()
    const ligatureShaper: NativeTextShaper = {
      ...base,
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result || request.run.text !== 'AB') return result
        return {
          ...result,
          glyphs: result.glyphs.map((glyph) => ({ ...glyph, clusterIndex: 0 })),
          clusters: [{ startUtf16: 0, endUtf16: 2, glyphStart: 0, glyphEnd: result.glyphs.length, advanceInlineMilliPoints: 2_000, unsafeToBreak: true }],
        }
      },
    }
    const element = nativeTextElement('overfull-cluster', 'AB', nativeTextBody(), { x: 0, y: 0, cx: 20_000, cy: 200_000 })
    const tree = await compileNativePptxSlide(authoredDeck([element]), 0, { textLayout: textLayout(ligatureShaper) })
    expect(findNode(tree, 'text', element.id).textBody).toMatchObject({ status: 'refused', paragraphs: [], refusalLabel: 'Exact text layout unavailable' })
    expect(tree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.wrapUnavailable', severity: 'refusal', elementId: element.id })]))
    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(tree, recording)
    expect(recording.finish()).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: 'placeholder', reason: 'textRefusal', sourceElementId: element.id, rect: { x: 0, y: 0, cx: 20_000, cy: 200_000 } }),
    ]))
  })

  it('handles provider-sized no-wrap and one-line square cluster vectors without per-cluster wrap materialization', async () => {
    const clusterCount = 150_000
    const largeClusterShaper: NativeTextShaper = {
      providerId: 'large-cluster-fixture', providerRevision: '1',
      shape({ run, startUtf16, endUtf16, font }) {
        return {
          startUtf16, endUtf16, face: font.face, glyphs: [],
          clusters: Array.from({ length: clusterCount }, (_, index) => ({
            startUtf16: index, endUtf16: index + 1, glyphStart: 0, glyphEnd: 0,
            advanceInlineMilliPoints: 0,
          })),
          metrics: { fontSizeMilliPoints: run.fontSizeMilliPoints, ascentMilliPoints: 8_000, descentMilliPoints: -2_000, lineGapMilliPoints: 2_000, lineHeightMilliPoints: 12_000 },
          advanceInlineMilliPoints: 0, advanceBlockMilliPoints: 0,
        }
      },
    }
    const elements = [
      nativeTextElement('large-cluster-no-wrap', 'A'.repeat(clusterCount), nativeTextBody({ wrap: 'none' })),
      nativeTextElement('large-cluster-square', 'A'.repeat(clusterCount), nativeTextBody()),
    ]
    const tree = await compileNativePptxSlide(authoredDeck(elements), 0, { textLayout: textLayout(largeClusterShaper) })
    for (const element of elements) {
      const body = findNode(tree, 'text', element.id).textBody
      expect(body.paragraphs).toHaveLength(1)
      const clusters = body.paragraphs[0]!.runs[0]!.clusters
      expect(clusters).toHaveLength(clusterCount)
      expect(clusters[0]).toMatchObject({ startUtf16: 0, endUtf16: 1, glyphStart: 0, glyphEnd: 0 })
      expect(clusters.at(-1)).toMatchObject({ startUtf16: clusterCount - 1, endUtf16: clusterCount })
    }
  })

  it('refuses incomplete cluster coverage, invalid glyph ownership, and negative cluster advances', async () => {
    const mutations: Array<[(segment: ShapedSegment) => ShapedSegment, string]> = [
      [(segment) => ({ ...segment, clusters: segment.clusters.slice(1) }), 'render.invalidShaping'],
      [(segment) => ({ ...segment, glyphs: segment.glyphs.map((glyph, index) => index === 0 ? { ...glyph, clusterIndex: 1 } : glyph) }), 'render.invalidShaping'],
      [(segment) => ({ ...segment, clusters: segment.clusters.map((cluster, index) => index === 0 ? { ...cluster, advanceInlineMilliPoints: -1 } : cluster) }), 'render.invalidProviderOutput'],
      [(segment) => ({ ...segment, metrics: { ...segment.metrics, lineHeightMilliPoints: segment.metrics.lineHeightMilliPoints - 1 } }), 'render.invalidShaping'],
    ]
    for (const [mutate, code] of mutations) {
      const base = fixtureShaper()
      const corrupt: NativeTextShaper = {
        providerId: 'corrupt-fixture', providerRevision: '1',
        async shape(request) {
          const result = await base.shape(request)
          if ('status' in result) return result
          return mutate(result)
        },
      }
      const tree = await compileNativePptxSlide(authoredDeck([textElement('invalid-clusters', 'left', 0)]), 0, { textLayout: textLayout(corrupt) })
      expect(findNode(tree, 'text', 'invalid-clusters').textBody).toMatchObject({ status: 'refused', paragraphs: [] })
      expect(tree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.refused', elementId: 'invalid-clusters' })]))
      expect(code).toMatch(/^render\./)
    }
  })

  it('accepts complete descending RTL clusters and glyphless zero-advance clusters, but refuses vertical layout', async () => {
    const base = fixtureShaper()
    const invisible: NativeTextShaper = {
      providerId: 'invisible-fixture', providerRevision: '1',
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result || request.run.text !== 'AB') return result
        return {
          ...result,
          glyphs: [{ ...result.glyphs[1]!, clusterIndex: 1 }],
          clusters: [
            { ...result.clusters[0]!, glyphStart: 0, glyphEnd: 0, advanceInlineMilliPoints: 0 },
            { ...result.clusters[1]!, glyphStart: 0, glyphEnd: 1 },
          ],
          advanceInlineMilliPoints: 1_000,
        }
      },
    }
    const invisibleTree = await compileNativePptxSlide(authoredDeck([textElement('invisible-cluster', 'left', 0)]), 0, { textLayout: textLayout(invisible) })
    expect(findNode(invisibleTree, 'text', 'invisible-cluster').textBody.paragraphs[0]!.runs[0]).toMatchObject({
      status: 'shaped', glyphs: [expect.objectContaining({ clusterIndex: 1 })],
      clusters: [expect.objectContaining({ glyphStart: 0, glyphEnd: 0, advanceInlineEmu: 0 }), expect.objectContaining({ glyphStart: 0, glyphEnd: 1 })],
    })

    let providerInvoked = false
    const verticalLayout = textLayout({
      ...fixtureShaper(),
      shape(request) { providerInvoked = true; return fixtureShaper().shape(request) },
    }, () => ({ direction: 'ttb', script: 'Hani', language: 'zh' }))
    const verticalTree = await compileNativePptxSlide(authoredDeck([textElement('vertical', 'left', 0)]), 0, { textLayout: verticalLayout })
    expect(providerInvoked).toBe(false)
    expect(findNode(verticalTree, 'text', 'vertical').textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    expect(verticalTree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.verticalUnsupported', severity: 'refusal' })]))

    const nativeVertical = nativeTextElement('native-vertical', 'AB', nativeTextBody())
    const nativeVerticalTree = await compileNativePptxSlide(authoredDeck([nativeVertical]), 0, { textLayout: verticalLayout })
    expect(providerInvoked).toBe(false)
    expect(findNode(nativeVerticalTree, 'text', nativeVertical.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    expect(nativeVerticalTree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.verticalUnsupported', severity: 'refusal', elementId: nativeVertical.id })]))
  })

  it('turns shaping refusals and missing chart previews into visible commands and diagnostics', async () => {
    const refused = await compileNativePptxSlide(authoredDeck([textElement('refused-text', 'left', 0)]), 0, { textLayout: textLayout(fixtureShaper((text) => text === 'CD')) })
    expect(findNode(refused, 'text', 'refused-text').textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    expect(refused.diagnostics.map((diagnostic) => diagnostic.code)).toContain('text.refused')
    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(refused, recording)
    expect(recording.finish()).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'placeholder', reason: 'textRefusal', sourceElementId: 'refused-text' })]))

    const noPreview = structuredClone(parsedFull)
    const chart = noPreview.slides[0]!.elements.find((element) => element.kind === 'chart')!
    if (chart.kind !== 'chart') throw new Error('chart fixture changed')
    delete chart.chart.previewAssetId
    const chartTree = await compileNativePptxSlide(noPreview, 0, { textLayout: textLayout() })
    expect(findNode(chartTree, 'placeholder', 'el-chart')).toMatchObject({ reason: 'missingPreview' })
    expect(chartTree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'chart.missingPreview', severity: 'refusal' })]))
  })

  it('records headlessly and replays through a host-owned context adapter', async () => {
    const tree = await compileNativePptxSlide(parsedFull, 0, { textLayout: textLayout() })
    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(tree, recording)
    const commands = recording.finish()
    expect(Object.isFrozen(commands)).toBe(true)
    expect(Object.isFrozen(commands[0])).toBe(true)
    expect(commands[0]).toMatchObject({ kind: 'beginSlide', background: 'FFFFFF' })
    expect(commands.at(-1)).toEqual({ kind: 'endSlide' })
    expect(commands).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: 'path', sourceElementId: 'el-line', tailArrow: true }),
      expect.objectContaining({ kind: 'image', role: 'chartPreview', assetId: 'a-preview' }),
      expect.objectContaining({ kind: 'glyphRun', sourceElementId: 'el-title' }),
    ]))
    const context: PaintCommand[] = []
    paintSlideRenderTreeToCanvas2D(tree, context, { execute(target, command) { target.push(command) } })
    expect(context).toEqual(commands)
  })

  it('enforces coordinate, depth, node, glyph, and paint-command budgets', async () => {
    await expect(compileNativePptxSlide(authoredDeck([textElement('bounded', 'left', 0)]), 0, { textLayout: textLayout(), maxNodes: 2 })).rejects.toMatchObject({ code: 'render.nodeBudget' })
    await expect(compileNativePptxSlide(authoredDeck([textElement('bounded', 'left', 0)]), 0, { textLayout: textLayout(), maxGlyphs: 1 })).rejects.toMatchObject({ code: 'render.glyphBudget' })
    await expect(compileNativePptxSlide(authoredDeck([textElement('bounded', 'left', 0)]), 0, { textLayout: textLayout(), maxClusters: 1 })).rejects.toMatchObject({ code: 'render.clusterBudget' })
    await expect(compileNativePptxSlide(authoredDeck([{ kind: 'shape', id: 'far', provenance: 'authored', transform: { x: 2_000_000, y: 0, cx: 1, cy: 1 }, preset: 'rect', paragraphs: [], passthrough: [], compatibility: { status: 'editable', diagnostics: [] } }]), 0, { textLayout: textLayout(), maxCoordinateEmu: 1_000_000 })).rejects.toMatchObject({ code: 'render.coordinateBudget' })
    const strokeDeck = authoredDeck([{ kind: 'shape', id: 'wide-stroke', provenance: 'authored', transform: { x: 0, y: 0, cx: 1, cy: 1 }, preset: 'rect', stroke: { color: '000000', widthEmu: 12_700 }, paragraphs: [], passthrough: [], compatibility: { status: 'editable', diagnostics: [] } }])
    strokeDeck.size = { cx: 10_000, cy: 10_000 }
    await expect(compileNativePptxSlide(strokeDeck, 0, { textLayout: textLayout(), maxCoordinateEmu: 10_000 })).rejects.toMatchObject({ code: 'render.coordinateBudget' })
    const nested: NativeElement = {
      kind: 'group', id: 'outer', provenance: 'authored', transform: { x: 0, y: 0, cx: 100, cy: 100 }, passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
      children: [{
        kind: 'group', id: 'inner', provenance: 'authored', transform: { x: 0, y: 0, cx: 100, cy: 100 }, passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
        children: [{ kind: 'shape', id: 'leaf', provenance: 'authored', transform: { x: 0, y: 0, cx: 10, cy: 10 }, preset: 'rect', paragraphs: [], passthrough: [], compatibility: { status: 'editable', diagnostics: [] } }],
      }],
    }
    await expect(compileNativePptxSlide(authoredDeck([nested]), 0, { textLayout: textLayout(), maxDepth: 1 })).rejects.toMatchObject({ code: 'render.depthBudget' })
    const recording = createRecordingPaintSurface(1)
    recording.push({ kind: 'endSlide' })
    expect(() => recording.push({ kind: 'endSlide' })).toThrowError(RenderCompileError)
    const tree = await compileNativePptxSlide(authoredDeck([]), 0, { textLayout: textLayout() })
    const atomic = createRecordingPaintSurface()
    expect(() => paintSlideRenderTree(tree, atomic, 1)).toThrowError(expect.objectContaining({ code: 'render.paintBudget' }))
    expect(atomic.commands).toEqual([])
  })

  it('emits deterministic bounded paths for every supported preset', () => {
    for (const preset of ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'rightArrow', 'pentagon', 'hexagon', 'star5'] as const) {
      const first = presetPath(preset, 1_000_003, 700_001)
      expect(presetPath(preset, 1_000_003, 700_001)).toEqual(first)
      expect(first.length).toBeLessThanOrEqual(11)
      expect(JSON.stringify(first)).not.toMatch(/\.\d/)
    }
  })

  it('uses DrawingML default pentagon guides instead of an inscribed polygon', () => {
    expect(presetPath('pentagon', 1_000_000, 1_000_000)).toEqual([
      {kind:'moveTo',x:1,y:381965},{kind:'lineTo',x:500000,y:0},
      {kind:'lineTo',x:999999,y:381965},{kind:'lineTo',x:809016,y:999997},
      {kind:'lineTo',x:190984,y:999997},{kind:'close'},
    ])
    expect(presetPath('pentagon', 1417740, 1317072)).toEqual([
      {kind:'moveTo',x:1,y:503075},{kind:'lineTo',x:708870,y:0},
      {kind:'lineTo',x:1417739,y:503075},{kind:'lineTo',x:1146975,y:1317069},
      {kind:'lineTo',x:270765,y:1317069},{kind:'close'},
    ])
  })

  it('paints preserved geometry with explicit omitted-text diagnostics and no invented glyphs', async () => {
    const shape:NativeElement={kind:'shape',id:'partial-pentagon',provenance:'authored',
      transform:{x:100000,y:100000,cx:1000000,cy:800000,quarterTurns:2},preset:'pentagon',fill:'123456',paragraphs:[],passthrough:[],
      compatibility:{status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.autoshape-text-layout-unavailable',message:'Vertical text omitted'}]}}
    const deck=authoredDeck([shape]);deck.compatibility=shape.compatibility;deck.slides[0]!.compatibility=shape.compatibility
    const tree=await compileNativePptxSlide(deck,0,{textLayout:textLayout()})
    const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
    const commands=surface.finish()
    expect(commands).toContainEqual(expect.objectContaining({kind:'path',fill:'123456'}))
    expect(commands.some(command=>command.kind==='glyphRun'||command.kind==='placeholder')).toBe(false)
    expect(tree.diagnostics).toContainEqual(expect.objectContaining({sourceCode:'pptx.autoshape-text-layout-unavailable'}))
  })

  it('preserves native AutoShape stroke semantics and placeholders refused custom geometry', async () => {
    const exact: NativeElement = {
      kind: 'shape', id: 'native-autoshape', provenance: 'authored',
      transform: { x: 100, y: 200, cx: 1_000_000, cy: 500_000 }, preset: 'rect', fill: 'DDEEFF',
      stroke: { color: '112233', widthEmu: 12_700, cap: 'square', join: 'miter', dash: 'solid', miterLimit: 800_000 },
      paragraphs: [], passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    }
    const tree = await compileNativePptxSlide(authoredDeck([exact]), 0, { textLayout: textLayout() })
    expect(findNode(tree, 'shape', exact.id)).toMatchObject({
      fill: { color: 'DDEEFF' },
      stroke: { color: '112233', widthEmu: 12_700, cap: 'square', join: 'miter', dash: 'solid', miterLimit: 800_000 },
      path: [{ kind: 'rect', rect: { x: 0, y: 0, cx: 1_000_000, cy: 500_000 } }],
    })
    const recording = createRecordingPaintSurface()
    paintSlideRenderTree(tree, recording)
    expect(recording.finish()).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: 'path', sourceElementId: exact.id, stroke: expect.objectContaining({ cap: 'square', join: 'miter', dash: 'solid', miterLimit: 800_000 }) }),
    ]))

    const refused: NativeElement = {
      kind: 'shape', id: 'custom-geometry', provenance: 'authored',
      transform: { x: 0, y: 0, cx: 100_000, cy: 100_000 }, paragraphs: [], passthrough: [],
      compatibility: { status: 'refused', diagnostics: [{ severity: 'refusal', code: 'pptx.autoshape-geometry-unavailable', message: 'custom geometry is opaque' }] },
    }
    const refusedDeck = authoredDeck([refused])
    refusedDeck.slides[0]!.compatibility = { status: 'refused', diagnostics: [{ severity: 'refusal', code: 'pptx.autoshape-geometry-unavailable', message: 'contains refused geometry', scope: { slideId: refusedDeck.slides[0]!.id, elementId: refused.id } }] }
    refusedDeck.compatibility = { status: 'refused', diagnostics: [{ severity: 'refusal', code: 'deck.refused', message: 'contains refused geometry' }] }
    const refusedTree = await compileNativePptxSlide(refusedDeck, 0, { textLayout: textLayout() })
    expect(findNode(refusedTree, 'placeholder', refused.id)).toMatchObject({ reason: 'refused', label: 'Unsupported shape' })

    refusedDeck.slides[0]!.compatibility = { status: 'refused', diagnostics: [{ severity: 'refusal', code: 'slide.markup-unavailable', message: 'slide root is opaque', scope: { slideId: refusedDeck.slides[0]!.id } }] }
    const refusedSlideTree = await compileNativePptxSlide(refusedDeck, 0, { textLayout: textLayout() })
    expect(refusedSlideTree.nodes).toEqual([expect.objectContaining({ kind: 'placeholder', sourceElementId: refusedDeck.slides[0]!.id, label: 'Slide rendering refused' })])
  })

  it('normalizes provider output exactly and fails closed on hostile runtime values', async () => {
    const cycle: Record<string, unknown> = { code: 'provider-failure', message: 'cycle', recoverable: false }
    cycle.cycle = cycle
    const hostileResolvers: NativeFontResolver[] = [
      { ...resolver, resolve() { return { status: 'unexpected' } as never } },
      { ...resolver, resolve({ run }) { return { status: 'resolved', face: face(run.font.weight), attemptedFaceIds: ['bad face id!'], decisions: [] } as never } },
      { ...resolver, resolve({ run }) { return { status: 'resolved', face: { ...face(run.font.weight), injected: true }, attemptedFaceIds: [], decisions: [] } as never } },
      { ...resolver, resolve({ run }) { return { status: 'resolved', face: face(run.font.weight), attemptedFaceIds: [], decisions: [cycle] } as never } },
      { ...resolver, load() { return { status: 'broken' } as never } },
    ]
    for (const hostile of hostileResolvers) {
      const tree = await compileNativePptxSlide(authoredDeck([textElement('hostile-provider', 'left', 0)]), 0, {
        textLayout: { ...textLayout(), resolver: hostile },
      })
      expect(findNode(tree, 'text', 'hostile-provider').textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    }

    const base = fixtureShaper()
    const hostileShaper: NativeTextShaper = {
      ...base,
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result) return result
        return { ...result, glyphs: result.glyphs.map((glyph, index) => index === 0 ? { ...glyph, injected: { cycle } } : glyph) } as never
      },
    }
    const hostileTree = await compileNativePptxSlide(authoredDeck([textElement('hostile-shaper', 'left', 0)]), 0, { textLayout: textLayout(hostileShaper) })
    expect(findNode(hostileTree, 'text', 'hostile-shaper').textBody).toMatchObject({ status: 'refused', paragraphs: [] })

    const mismatchedLoad: NativeFontResolver = {
      ...resolver,
      load(resolved) {
        return { face: { ...resolved, family: 'Changed Family' }, bytes: new Uint8Array([1]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 0 } }
      },
    }
    const mismatchedTree = await compileNativePptxSlide(authoredDeck([textElement('mismatched-face', 'left', 0)]), 0, { textLayout: { ...textLayout(), resolver: mismatchedLoad } })
    expect(findNode(mismatchedTree, 'text', 'mismatched-face').textBody).toMatchObject({ status: 'refused', paragraphs: [] })
  })

  it('binds shaping to manifest bytes and design metrics, caches one bounded font load, and refuses provider mutation atomically', async () => {
    const twoRuns = nativeTextElement('provider-authority', 'unused', nativeTextBody({ wrap: 'none' }))
    twoRuns.paragraphs = [{
      align: 'left', level: 0, bullet: false,
      runs: [
        { text: 'A', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 },
        { text: 'B', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 },
      ],
    }]
    let loads = 0
    const countingResolver: NativeFontResolver = {
      ...resolver,
      load(resolved) { loads++; return resolver.load(resolved) },
    }
    const exactTree = await compileNativePptxSlide(authoredDeck([twoRuns]), 0, { textLayout: { ...textLayout(), resolver: countingResolver } })
    expect(findNode(exactTree, 'text', twoRuns.id).textBody).toMatchObject({ status: 'laidOut' })
    expect(loads).toBe(1)

    let byteMutationShapeCalls = 0
    const mutatingShaper: NativeTextShaper = {
      ...fixtureShaper(),
      async shape(request) {
        byteMutationShapeCalls++
        const result = await fixtureShaper().shape(request)
        request.font.bytes[0] = request.font.bytes[0]! ^ 0xff
        return result
      },
    }
    const byteMutationTree = await compileNativePptxSlide(authoredDeck([twoRuns]), 0, { textLayout: textLayout(mutatingShaper) })
    expect(byteMutationShapeCalls).toBe(1)
    expect(findNode(byteMutationTree, 'text', twoRuns.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })

    let shapedAfterBadDigest = false
    const badDigestResolver: NativeFontResolver = {
      ...resolver,
      load(resolved) { return { face: resolved, bytes: new Uint8Array([9]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 } } },
    }
    const guardedShaper: NativeTextShaper = { ...fixtureShaper(), shape(request) { shapedAfterBadDigest = true; return fixtureShaper().shape(request) } }
    const badDigestTree = await compileNativePptxSlide(authoredDeck([twoRuns]), 0, { textLayout: { ...textLayout(guardedShaper), resolver: badDigestResolver } })
    expect(shapedAfterBadDigest).toBe(false)
    expect(findNode(badDigestTree, 'text', twoRuns.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })

    const digestlessManifest: NativeFontManifest = {
      ...manifest,
      faces: manifest.faces.map((item) => ({ ...item, source: { ...item.source, contentDigest: undefined } })),
    }
    const digestlessTree = await compileNativePptxSlide(authoredDeck([twoRuns]), 0, {
      textLayout: { ...textLayout(), manifest: digestlessManifest },
    })
    expect(findNode(digestlessTree, 'text', twoRuns.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })

    const metricForgery: NativeTextShaper = {
      ...fixtureShaper(),
      async shape(request) {
        const result = await fixtureShaper().shape(request)
        if ('status' in result) return result
        return { ...result, metrics: { ...result.metrics, ascentMilliPoints: result.metrics.ascentMilliPoints + 1, lineHeightMilliPoints: result.metrics.lineHeightMilliPoints + 1 } }
      },
    }
    const metricTree = await compileNativePptxSlide(authoredDeck([twoRuns]), 0, { textLayout: textLayout(metricForgery) })
    expect(findNode(metricTree, 'text', twoRuns.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })

    const identityResolver: NativeFontResolver = {
      ...resolver,
      resolve(request) {
        ;(identityResolver as { providerRevision: string }).providerRevision = 'forged'
        return resolver.resolve(request)
      },
    }
    const identityTree = await compileNativePptxSlide(authoredDeck([twoRuns]), 0, { textLayout: { ...textLayout(), resolver: identityResolver } })
    expect(findNode(identityTree, 'text', twoRuns.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })

    const runMutationResolver: NativeFontResolver = {
      ...resolver,
      resolve(request) {
        ;(request.run as { text: string }).text = 'forged'
        return resolver.resolve(request)
      },
    }
    const runMutationTree = await compileNativePptxSlide(authoredDeck([twoRuns]), 0, { textLayout: { ...textLayout(), resolver: runMutationResolver } })
    expect(findNode(runMutationTree, 'text', twoRuns.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })

    const mutableDeck = authoredDeck([structuredClone(twoRuns)])
    const deckMutationResolver: NativeFontResolver = {
      ...resolver,
      resolve(request) {
        ;(mutableDeck.slides[0]!.elements[0] as typeof twoRuns).paragraphs[0]!.runs[1]!.text = 'FORGED'
        return resolver.resolve(request)
      },
    }
    const snapshottedTree = await compileNativePptxSlide(mutableDeck, 0, { textLayout: { ...textLayout(), resolver: deckMutationResolver } })
    const snapshottedText = findNode(snapshottedTree, 'text', twoRuns.id).textBody.paragraphs
      .flatMap((line) => line.runs.map((run) => run.text)).join('')
    expect(snapshottedText).toBe('AB')
  })

  it('refuses paragraph semantics and separator metrics that cannot be painted exactly without emitting partial runs', async () => {
    const semanticCases = [
      { align: undefined, level: 0, bullet: false },
      { align: 'left' as const, level: 1, bullet: false },
      { align: 'left' as const, level: 0, bullet: true },
    ]
    let providerCalls = 0
    const countingShaper: NativeTextShaper = { ...fixtureShaper(), shape(request) { providerCalls++; return fixtureShaper().shape(request) } }
    for (const [index, semantics] of semanticCases.entries()) {
      const element = nativeTextElement(`paragraph-semantics-${index}`, 'AB', nativeTextBody())
      Object.assign(element.paragraphs[0]!, semantics)
      if (semantics.align === undefined) delete element.paragraphs[0]!.align
      const tree = await compileNativePptxSlide(authoredDeck([element]), 0, { textLayout: textLayout(countingShaper) })
      expect(findNode(tree, 'text', element.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
      expect(tree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.paragraphSemanticsUnavailable', elementId: element.id })]))
    }
    expect(providerCalls).toBe(0)

    const emptyParagraph = nativeTextElement('empty-paragraph-metrics', 'AB', nativeTextBody())
    emptyParagraph.paragraphs[0]!.runs = []
    const emptyParagraphTree = await compileNativePptxSlide(authoredDeck([emptyParagraph]), 0, { textLayout: textLayout(countingShaper) })
    expect(findNode(emptyParagraphTree, 'text', emptyParagraph.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    expect(emptyParagraphTree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.metricsUnavailable', elementId: emptyParagraph.id })]))
    expect(providerCalls).toBe(0)

    const inherited = nativeTextElement('unresolved-inheritance', 'AB', nativeTextBody())
    inherited.paragraphs[0]!.runs[0] = { text: 'AB' }
    const inheritedDeck = authoredDeck([inherited])
    inheritedDeck.slides[0]!.compatibility = {
      status: 'preserveOnly',
      diagnostics: [{ severity: 'warning', code: 'pptx.unsupported-master-dependency', message: 'master text styles are preserved but unresolved' }],
    }
    inheritedDeck.compatibility = {
      status: 'preserveOnly',
      diagnostics: [{ severity: 'warning', code: 'pptx.unsupported-master-dependency', message: 'master text styles are preserved but unresolved' }],
    }
    const inheritedTree = await compileNativePptxSlide(inheritedDeck, 0, { textLayout: textLayout(countingShaper) })
    expect(findNode(inheritedTree, 'text', inherited.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    expect(inheritedTree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.inheritanceUnavailable', elementId: inherited.id })]))
    expect(providerCalls).toBe(0)

    const token = nativeTextElement('theme-token', 'AB', nativeTextBody())
    token.paragraphs[0]!.runs[0] = { text: 'AB', fontFamily: '+mj-lt', fontSizeHundredthPt: 1_000 }
    const tokenTree = await compileNativePptxSlide(authoredDeck([token]), 0, { textLayout: textLayout(countingShaper) })
    expect(findNode(tokenTree, 'text', token.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
    expect(tokenTree.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'text.inheritanceUnavailable', elementId: token.id })]))
    expect(providerCalls).toBe(0)

    const separator = nativeTextElement('separator-metrics', 'unused', nativeTextBody(), { x: 0, y: 0, cx: 25_400, cy: 500_000 })
    separator.paragraphs = [{ align: 'left', level: 0, bullet: false, runs: [
      { text: 'A', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 },
      { text: ' ', fontFamily: 'Aptos', fontSizeHundredthPt: 2_000 },
      { text: 'B', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 },
    ] }]
    const separatorTree = await compileNativePptxSlide(authoredDeck([separator]), 0, { textLayout: textLayout() })
    expect(findNode(separatorTree, 'text', separator.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
  })

  it('lays out self-contained native text while unresolved theme/master parts remain preserve-only', async () => {
    const element = nativeTextElement('self-contained-with-theme', 'AB', nativeTextBody())
    const deck = authoredDeck([element])
    const diagnostics = [
      { severity: 'warning' as const, code: 'pptx.unsupported-theme-dependency', message: 'fmtScheme remains opaque' },
      { severity: 'warning' as const, code: 'pptx.unsupported-master-dependency', message: 'master text styles are preserved but unresolved' },
      { severity: 'warning' as const, code: 'pptx.unsupported-layout-dependency', message: 'layout placeholders are preserved but unresolved' },
    ]
    deck.compatibility = { status: 'preserveOnly', diagnostics }
    deck.slides[0]!.compatibility = { status: 'preserveOnly', diagnostics }
    const tree = await compileNativePptxSlide(deck, 0, { textLayout: textLayout() })
    const textBody = findNode(tree, 'text', element.id).textBody
    expect(textBody).toMatchObject({ fidelity: 'native', status: 'laidOut' })
    expect(textBody.paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join('')).toBe('AB')
    expect(tree.diagnostics.map((diagnostic) => diagnostic.code)).not.toContain('text.inheritanceUnavailable')
  })

  it('rejects positive-advance or painted default-ignorable clusters from a hostile shaper', async () => {
    const base = fixtureShaper()
    const hostile: NativeTextShaper = {
      ...base,
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result) return result
        const control = result.clusters.findIndex((cluster) => request.run.text.charCodeAt(cluster.startUtf16) === 0x200b)
        if (control < 0) return result
        const clusters = result.clusters.map((cluster, index) => index === control ? { ...cluster, glyphStart: 1, glyphEnd: 2, advanceInlineMilliPoints: 1_000 } : cluster)
        const glyphs = [...result.glyphs, { glyphId: 1, clusterIndex: control, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 }]
        return { ...result, glyphs, clusters, advanceInlineMilliPoints: result.advanceInlineMilliPoints + 1_000 }
      },
    }
    const element = nativeTextElement('hostile-zwsp', 'A\u200bB', nativeTextBody({ wrap: 'none' }))
    const tree = await compileNativePptxSlide(authoredDeck([element]), 0, { textLayout: textLayout(hostile) })
    expect(findNode(tree, 'text', element.id).textBody).toMatchObject({ status: 'refused', paragraphs: [] })
  })

  it('preserves only valid optional provider fields after normalization', async () => {
    const decisionResolver: NativeFontResolver = {
      ...resolver,
      resolve({ run }) {
        return {
          status: 'resolved', face: face(run.font.weight, run.font.families[0]), attemptedFaceIds: ['fixture.regular'],
          decisions: [{ code: 'font-not-found', message: 'fallback inspected', recoverable: true, faceId: 'fixture.regular', startUtf16: 0, endUtf16: 1 }],
        }
      },
    }
    const base = fixtureShaper()
    const flagShaper: NativeTextShaper = {
      ...base,
      async shape(request) {
        const result = await base.shape(request)
        if ('status' in result) return result
        return {
          ...result,
          clusters: result.clusters.map((cluster, index) => index === 0 ? { ...cluster, unsafeToBreak: true, whitespace: false } : cluster),
        }
      },
    }
    const tree = await compileNativePptxSlide(authoredDeck([textElement('normalized-provider', 'left', 0)]), 0, {
      textLayout: { ...textLayout(flagShaper), resolver: decisionResolver },
    })
    const run = findNode(tree, 'text', 'normalized-provider').textBody.paragraphs[0]!.runs[0]!
    expect(run.attemptedFaceIds).toEqual(['fixture.regular'])
    expect(run.decisions).toEqual([{ code: 'font-not-found', message: 'fallback inspected', recoverable: true, faceId: 'fixture.regular', startUtf16: 0, endUtf16: 1 }])
    expect(run.clusters[0]).toMatchObject({ unsafeToBreak: true, whitespace: false })
    expect(Object.keys(run.clusters[0]!).sort()).toEqual(['advanceInlineEmu', 'advanceInlineMilliPoints', 'endUtf16', 'glyphEnd', 'glyphStart', 'startUtf16', 'unsafeToBreak', 'whitespace'])
  })

  it('rejects invalid native input before any provider is called', async () => {
    const invalid = authoredDeck([]) as NativePptxDeck & { unexpected?: boolean }
    invalid.unexpected = true
    let invoked = false
    const layout = textLayout()
    const guarded: NativePptxTextLayout = { ...layout, resolver: { ...layout.resolver, resolve(request) { invoked = true; return layout.resolver.resolve(request) } } }
    await expect(compileNativePptxSlide(invalid, 0, { textLayout: guarded })).rejects.toThrow(/invalid native PPTX contract/)
    expect(invoked).toBe(false)
  })
})

describe('DOM-free dependency guard', () => {
  it('keeps the runtime source free of UI frameworks, browser globals, JSX, and layout markup', () => {
    const sourceRoot = resolve(import.meta.dirname)
    const files = readdirSync(sourceRoot).filter((name) => name.endsWith('.ts') && !name.endsWith('.test.ts'))
    const source = files.map((name) => readFileSync(resolve(sourceRoot, name), 'utf8')).join('\n')
    expect(source).not.toMatch(/from\s+['"](?:react|konva|react-konva)['"]/)
    expect(source).not.toMatch(/\b(?:window|document|OffscreenCanvas|CanvasRenderingContext2D|devicePixelRatio)\b\s*[.[]/)
    expect(source).not.toMatch(/\b(?:measureText|getContext)\b/)
    expect(source).not.toMatch(/Math\.(?:min|max)\(\s*\.\.\./)
    expect(source).not.toMatch(/function\s+wrapAtoms\b/)
    expect(source).not.toMatch(/\.text\.slice\(\s*cluster\.(?:startUtf16|endUtf16)/)
    expect(source).not.toMatch(/for\s*\([^)]*=\s*0;[^)]*<\s*(?:clusterEnd|glyphStart)\b/)
    expect(source).not.toMatch(/<svg\b|<canvas\b|\.tsx\b/)
    expect(source).not.toMatch(/\{\s*\.\.\.(?:glyph|cluster|decision|face|resolution|loaded|shaped)\b/)
    const packageJson = JSON.parse(readFileSync(resolve(sourceRoot, '../package.json'), 'utf8')) as { dependencies?: Record<string, string> }
    expect(Object.keys(packageJson.dependencies ?? {}).sort()).toEqual(['@injoffice/font-metrics', '@injoffice/pptx-native', '@noble/hashes'].sort())
  })
})
