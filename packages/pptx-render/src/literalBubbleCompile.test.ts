import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it} from 'vitest'
import type {NativePptxDeck,NativeLiteralBubble} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,type NativePptxTextLayout} from './index.js'
const layout:NativePptxTextLayout={
 manifest:{version:1,manifestId:'geometry-only',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused',contentDigest:`sha256:${'0'.repeat(64)}`}}],fallbackChains:[]},
 resolver:{providerId:'unused',providerRevision:'1',resolve(){throw Error('No text')},load(){throw Error('No font')}},shaper:{providerId:'unused',providerRevision:'1',shape(){throw Error('No glyphs')}},defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'},
}

function fixture(){
 const deck=JSON.parse(readFileSync(resolve(import.meta.dirname,'../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')) as NativePptxDeck
 const chart=deck.slides[0]!.elements.find(e=>e.kind==='chart')!;if(chart.kind!=='chart')throw Error('missing chart')
 deck.slides[0]!.elements=[chart];chart.transform={x:0,y:0,cx:1000,cy:1000}
 const bubble:NativeLiteralBubble={profile:'literal-bubble-v1',dataOrigin:'literal',bubbleScale:100,sizeRepresents:'area',series:[{index:3,order:0,xValues:['1','0'],values:['0','0'],sizes:['4','1'],colors:['#123456','#ABCDEF']}],xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:false,color:'#000000',widthEmu:10,min:'-1',max:'1',crossesAt:'0'},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:false,color:'#000000',widthEmu:10,min:'-1',max:'1',crossesAt:'0'}}
 chart.chart.literalBubble=bubble;return {deck,chart,bubble}
}
it('compiles literal circles only after opt-in, clips complete ink and preserves source',async()=>{
 const {deck}=fixture(),before=JSON.stringify(deck)
 expect((await compileNativePptxSlide(deck,0,{textLayout:layout})).nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalBubblePreview:true}),surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const commands=surface.finish(),paths=commands.filter(c=>c.kind==='path')
 expect(paths).toHaveLength(4);expect(paths[0]!.path).toHaveLength(4);expect(paths[0]!.fill).toBe('#123456');expect(paths[0]!.path[0]).toEqual({kind:'moveTo',x:1100,y:500})
 expect(tree.nodes[0]!.kind==='group'&&tree.nodes[0].clip).toEqual({kind:'rect',rect:{x:0,y:0,cx:1000,cy:1000}})
 expect(tree.diagnostics.some(d=>d.code==='chart.bubblePreview')).toBe(true);expect(JSON.stringify(deck)).toBe(before)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalBubblePreview:true,maxCoordinateEmu:1000})).rejects.toThrow()
})
it('retains axes at source scale zero and all-zero sizes; width mode changes only radii',async()=>{
 const {deck,chart,bubble}=fixture()
 for(const zero of [{...bubble,bubbleScale:0},{...bubble,series:bubble.series.map(s=>({...s,sizes:['0','-0']}))}]){
  chart.chart.literalBubble=zero
  const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalBubblePreview:true}),surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
  expect(surface.finish().filter(c=>c.kind==='path')).toHaveLength(2)
 }
 chart.chart.literalBubble={...bubble,sizeRepresents:'w'}
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalBubblePreview:true}),surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 expect(surface.finish().filter(c=>c.kind==='path')[1]!.path[0]).toEqual({kind:'moveTo',x:525,y:500})
})
it('enforces aggregate node budgets for the maximum admitted 4096 circles',async()=>{
 const {deck,chart,bubble}=fixture()
 chart.chart.literalBubble={...bubble,series:Array.from({length:16},(_,order)=>({index:order,order,xValues:Array(256).fill('0'),values:Array(256).fill('0'),sizes:Array(256).fill('1'),colors:Array(256).fill('#123456')}))}
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalBubblePreview:true,maxNodes:4095})).rejects.toThrow()
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalBubblePreview:true,maxNodes:4100})
 expect(tree.nodes[0]!.kind==='group'&&tree.nodes[0].children.length).toBe(4098)
})
