import { describe, expect, it } from 'vitest'
import { compileWireDeckToNativeV1 } from '../../pptx-authored/src/index.js'
import type { NativePptxDeck, NativeTextBodyLayout } from '@injoffice/pptx-native'
import type { NativePptxTextLayout } from './types.js'
import { compileNativePptxSlide, createRecordingPaintSurface, paintSlideRenderTree, presetPath } from './index.js'
import { defaultPresetTextRect } from './geometry.js'

// Independent evaluation of the published OOXML guide angles (degrees here,
// not production radical formulas). Keep precision until final point rounding.
function reference(w: number, h: number) {
  const cos = (d: number) => Math.cos(d * Math.PI / 180)
  const sin = (d: number) => Math.sin(d * Math.PI / 180)
  const sw = w / 2 * 1.05146, sh = h / 2 * 1.10557, a = 19098 / 50000
  const x1 = w / 2 - sw * cos(18), x2 = w / 2 - sw * cos(306)
  const x3 = w / 2 + sw * cos(306), x4 = w / 2 + sw * cos(18)
  const y1 = sh - sh * sin(18), y2 = sh - sh * sin(306)
  const sx1 = w / 2 - sw * a * cos(342), sx2 = w / 2 - sw * a * cos(54)
  const sx3 = w / 2 + sw * a * cos(54), sx4 = w / 2 + sw * a * cos(342)
  const sy1 = sh - sh * a * sin(54), sy2 = sh - sh * a * sin(342), sy3 = sh + sh * a
  return {
    points: [[x1,y1],[sx2,sy1],[w/2,0],[sx3,sy1],[x4,y1],[sx4,sy2],[x3,y2],[w/2,sy3],[x2,y2],[sx1,sy2]],
    rect: { x:Math.round(sx1), y:Math.round(sy1), cx:Math.round(sx4)-Math.round(sx1), cy:Math.round(sy3)-Math.round(sy1) },
  }
}

const layout: NativePptxTextLayout = {
  manifest:{version:1,manifestId:'geometry-only',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused',contentDigest:`sha256:${'0'.repeat(64)}`}}],fallbackChains:[]},
  resolver:{providerId:'unused',providerRevision:'1',resolve(){throw new Error('No text requested')},load(){throw new Error('No font requested')}},
  shaper:{providerId:'unused',providerRevision:'1',shape(){throw new Error('No glyphs requested')}},
  defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'},
}
const body: NativeTextBodyLayout = {leftInsetEmu:1000,rightInsetEmu:2000,topInsetEmu:3000,bottomInsetEmu:4000,wrap:'square',verticalAnchor:'top',autoFit:'none',horizontalOverflow:'overflow',verticalOverflow:'overflow'}

function authoredStar(): NativePptxDeck {
  const result=compileWireDeckToNativeV1({slides:[{background:'#FFFFFF',shapes:[{kind:'star5',x:0,y:0,cx:2400000,cy:1200000,fill:'#336699'}]}]})
  if(!result.ok)throw new Error(JSON.stringify(result.issues))
  return structuredClone(result.deck)
}

describe('DrawingML default star5',()=>{
  it.each([[1000000,1000000],[4200000,2500000],[2500000,4200000],[1,2],[100000003,70000001]])('matches default guides for %sx%s frames',(w,h)=>{
    const ref=reference(w,h)
    expect(presetPath('star5',w,h)).toEqual([...ref.points.map(([x,y],i)=>({kind:i===0?'moveTo':'lineTo',x:Math.round(x!),y:Math.round(y!)})),{kind:'close'}])
    expect(defaultPresetTextRect('star5',w,h)).toEqual(ref.rect)
  })

  it('fills the intended outer frame and uses a compact internal text region',()=>{
    const points=presetPath('star5',1000000,1000000)
    expect(points[0]).toEqual({kind:'moveTo',x:1,y:381965})
    expect(points[6]).toEqual({kind:'lineTo',x:809016,y:999997})
    const rect=defaultPresetTextRect('star5',1000000,1000000)
    expect(rect.x).toBeGreaterThan(300000)
    expect(rect.y).toBeGreaterThan(380000)
    expect(rect.cx).toBeLessThan(400000)
  })

  it('uses the public authored compiler and preset text region through rotation and paint',async()=>{
    const deck=authoredStar(),star=deck.slides[0]!.elements[0]!
    if(star.kind!=='shape')throw new Error('Expected authored shape')
    star.textBody=body;star.transform.quarterTurns=1
    const before=JSON.stringify(deck),r=reference(star.transform.cx,star.transform.cy).rect
    const tree=await compileNativePptxSlide(deck,0,{textLayout:layout})
    const node=tree.nodes[0]!
    expect(node).toMatchObject({kind:'shape',preset:'star5',textBody:{status:'laidOut',bounds:{x:r.x+1000,y:r.y+3000,cx:r.cx-3000,cy:r.cy-7000}}})
    const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
    expect(surface.finish()).toContainEqual(expect.objectContaining({kind:'path',path:presetPath('star5',star.transform.cx,star.transform.cy)}))
    expect(JSON.stringify(deck)).toBe(before)
    star.textBody={...body,leftInsetEmu:Math.floor(star.transform.cx/3),rightInsetEmu:Math.floor(star.transform.cx/3)}
    await expect(compileNativePptxSlide(deck,0,{textLayout:layout})).rejects.toMatchObject({code:'render.coordinateBudget'})
  })

  it('rejects attempted authored adjustment metadata instead of ignoring it',()=>{
    const result=compileWireDeckToNativeV1({slides:[{background:'#FFFFFF',shapes:[{kind:'star5',x:0,y:0,cx:2400000,cy:1200000,adjustments:{adj:50000}}]}]})
    expect(result.ok).toBe(false)
  })
})
