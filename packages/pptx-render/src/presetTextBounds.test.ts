import { describe, expect, it } from 'vitest'
import { compileWireDeckToNativeV1 } from '../../pptx-authored/src/index.js'
import type { NativeFontManifest, NativeFontResolver, NativeTextRefusal, NativeTextShaper, ResolvedFontFace, ShapedGlyph, ShapedCluster } from '@injoffice/font-metrics/layout'
import type { NativePptxDeck, NativeShapePreset, NativeTextBodyLayout } from '@injoffice/pptx-native'
import { compileNativePptxSlide, createRecordingPaintSurface, paintSlideRenderTree, presetPath, type NativePptxTextLayout } from './index.js'
import { defaultPresetTextRect } from './geometry.js'

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

const body:NativeTextBodyLayout={leftInsetEmu:1000,rightInsetEmu:2000,topInsetEmu:3000,bottomInsetEmu:4000,wrap:'square',verticalAnchor:'top',autoFit:'none',horizontalOverflow:'overflow',verticalOverflow:'overflow'}
function deckFor(preset:NativeShapePreset):NativePptxDeck {
  const result=compileWireDeckToNativeV1({slides:[{background:'#FFFFFF',shapes:[{kind:preset,x:0,y:0,cx:2000000,cy:1000000,fill:'#336699',paragraphs:[{runs:[{text:'AB',font:'Fixture Sans',sizePt:12,color:'#112233'}]}]}]}]})
  if(!result.ok)throw new Error(JSON.stringify(result.issues))
  return structuredClone(result.deck)
}
const cases=[
  {preset:'ellipse' as const,bounds:{x:292893,y:146447,cx:1414214,cy:707106}},
  {preset:'triangle' as const,bounds:{x:500000,y:500000,cx:1000000,cy:500000}},
  {preset:'diamond' as const,bounds:{x:500000,y:250000,cx:1000000,cy:500000}},
]
describe('DrawingML primitive preset text bounds',()=>{
  it.each(cases)('$preset uses its normative default text rectangle',({preset,bounds})=>{
    expect(defaultPresetTextRect(preset,2000000,1000000)).toEqual(bounds)
    if(preset==='ellipse') expect(defaultPresetTextRect(preset,1000000,2000000)).toEqual({x:146447,y:292893,cx:707106,cy:1414214})
    else expect(defaultPresetTextRect(preset,1000000,2000000)).toEqual({x:250000,y:preset==='triangle'?1000000:500000,cx:500000,cy:1000000})
  })
  it('retains final-edge rounding for odd/tiny geometry',()=>{
    expect(defaultPresetTextRect('triangle',3,5)).toEqual({x:1,y:3,cx:1,cy:2})
    expect(defaultPresetTextRect('diamond',3,5)).toEqual({x:1,y:1,cx:1,cy:3})
    expect(defaultPresetTextRect('diamond',1,1)).toEqual({x:0,y:0,cx:1,cy:1})
    const w=100003,h=70001,dx=w/2*Math.cos(Math.PI/4),dy=h/2*Math.sin(Math.PI/4)
    expect(defaultPresetTextRect('ellipse',w,h)).toEqual({x:Math.round(w/2-dx),y:Math.round(h/2-dy),cx:Math.round(w/2+dx)-Math.round(w/2-dx),cy:Math.round(h/2+dy)-Math.round(h/2-dy)})
  })
  it.each(cases)('$preset lays out and paints actual glyphs after insets and rotation',async({preset,bounds})=>{
    for(const q of [0,1,2,3] as const){
      const deck=deckFor(preset),shape=deck.slides[0]!.elements[0]!
      if(shape.kind!=='shape')throw new Error('Expected authored shape')
      shape.textBody=body;if(q)shape.transform.quarterTurns=q
      const tree=await compileNativePptxSlide(deck,0,{textLayout:textLayout()})
      const node=tree.nodes[0]!
      if(node.kind!=='shape'||!node.textBody)throw new Error('Missing laid-out body')
      expect(node.textBody).toMatchObject({status:'laidOut',bounds:{x:bounds.x+1000,y:bounds.y+3000,cx:bounds.cx-3000,cy:bounds.cy-7000}})
      expect(node.textBody.paragraphs[0]!.runs[0]!.x).toBe(bounds.x+1000)
      expect(node.textBody.paragraphs[0]!.runs[0]!.baselineY).toBeGreaterThan(bounds.y+3000)
      expect(node.textBody.paragraphs[0]!.runs[0]!.glyphs).toHaveLength(2)
      expect(node.path).toEqual(presetPath(preset,2000000,1000000))
      const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
      expect(surface.finish().filter(c=>c.kind==='glyphRun')).toHaveLength(1)
      const before=JSON.stringify(shape.paragraphs)
      delete shape.textBody
      const legacy=await compileNativePptxSlide(deck,0,{textLayout:textLayout()})
      expect(legacy.nodes[0]).toMatchObject({textBody:{bounds:{x:0,y:0,cx:2000000,cy:1000000},fidelity:'legacyUnavailable'}})
      expect(JSON.stringify(shape.paragraphs)).toBe(before)
    }
  })
  it.each(cases)('$preset refuses insets that fit the frame but collapse its text region',async({preset})=>{
    const deck=deckFor(preset),shape=deck.slides[0]!.elements[0]!
    if(shape.kind!=='shape')throw new Error('Expected shape')
    shape.textBody={...body,leftInsetEmu:800000,rightInsetEmu:800000}
    await expect(compileNativePptxSlide(deck,0,{textLayout:textLayout()})).rejects.toMatchObject({code:'render.coordinateBudget'})
  })
  it('refuses a tiny triangle text region that rounds to zero height',async()=>{
    const deck=deckFor('triangle'),shape=deck.slides[0]!.elements[0]!
    if(shape.kind!=='shape')throw new Error('Expected shape')
    shape.transform={x:0,y:0,cx:1,cy:1}
    shape.textBody={...body,leftInsetEmu:0,rightInsetEmu:0,topInsetEmu:0,bottomInsetEmu:0}
    await expect(compileNativePptxSlide(deck,0,{textLayout:textLayout()})).rejects.toMatchObject({code:'render.coordinateBudget'})
  })

})
