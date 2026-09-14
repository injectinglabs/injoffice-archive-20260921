import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it} from 'vitest'
import {validateNativePptx,type NativePptxDeck,type NativeLiteralRadar} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,type NativePptxTextLayout} from './index.js'
const layout:NativePptxTextLayout={
 manifest:{version:1,manifestId:'geometry-only',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused',contentDigest:`sha256:${'0'.repeat(64)}`}}],fallbackChains:[]},
 resolver:{providerId:'unused',providerRevision:'1',resolve(){throw Error('No text')},load(){throw Error('No font')}},shaper:{providerId:'unused',providerRevision:'1',shape(){throw Error('No glyphs')}},defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'},
}
function fixture(){
 const deck=JSON.parse(readFileSync(resolve(import.meta.dirname,'../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')) as NativePptxDeck
 const chart=deck.slides[0]!.elements.find(e=>e.kind==='chart')!;if(chart.kind!=='chart')throw Error('chart')
 deck.slides[0]!.elements=[chart]
 const radar:NativeLiteralRadar={profile:'literal-radar-v1',dataOrigin:'literal',style:'standard',categories:['A','B','C','D'],series:[{index:3,order:1,color:'#123456',widthEmu:12700,values:['10','10','10','10']},{index:9,order:0,color:'#ABCDEF',widthEmu:12700,values:['5','5','5','5']}],categoryAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:true},valueAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:false,color:'#AA00AA',widthEmu:12700,min:'0',max:'10',crossesAt:'0'}}
 chart.chart.literalRadar=radar;chart.transform={x:0,y:0,cx:1000000,cy:1000000}
 return {deck,chart,radar}
}
it('keeps default off, series order, standard/filled spoke layering and original authority',async()=>{
 const {deck,chart,radar}=fixture(),before=JSON.stringify(deck)
 expect((await compileNativePptxSlide(deck,0,{textLayout:layout})).nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalRadarPreview:true}),paint=createRecordingPaintSurface();paintSlideRenderTree(tree,paint)
 expect(paint.finish().filter(n=>n.kind==='path').map(n=>n.kind==='path'?n.stroke?.color:undefined)).toEqual(['#AA00AA','#123456','#ABCDEF'])
 expect(JSON.stringify(deck)).toBe(before)
 chart.chart.literalRadar={...radar,style:'filled',series:radar.series.map(s=>({...s,fill:s.color}))}
 const filled=await compileNativePptxSlide(deck,0,{textLayout:layout,literalRadarPreview:true}),surface=createRecordingPaintSurface();paintSlideRenderTree(filled,surface)
 expect(surface.finish().filter(n=>n.kind==='path').map(n=>n.kind==='path'?n.stroke?.color:undefined)).toEqual(['#123456','#ABCDEF','#AA00AA'])
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalRadarPreview:true,maxNodes:2})).rejects.toThrow()
})
it('qualifies complete stroke+numerical ink before clipping for ordinary and source physical frames',async()=>{
 for(const physical of [false,true]){
  const {deck,chart}=fixture();deck.size={cx:2000000,cy:2000000};chart.transform={x:1000000,y:1000000,cx:1000000,cy:1000000}
  if(physical){chart.graphicFrameLayout='source-anchored-v1';chart.compatibility.status='preserveOnly'}
  await expect(compileNativePptxSlide(deck,0,{textLayout:layout,maxCoordinateEmu:2000002})).resolves.toBeDefined()
  await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalRadarPreview:true,maxCoordinateEmu:2000002})).rejects.toThrow()
 }
})
it('retains generic command limits at maximum radar inventory and refuses unqualified profiles',async()=>{
 const {deck,chart,radar}=fixture();chart.chart.literalRadar={...radar,categories:Array(256).fill('A'),series:Array.from({length:16},(_,i)=>({...radar.series[0]!,index:i,order:15-i,values:Array(256).fill('10')}))}
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalRadarPreview:true}),paint=createRecordingPaintSurface();paintSlideRenderTree(tree,paint)
 const paths=paint.finish().filter(n=>n.kind==='path');expect(paths).toHaveLength(17);expect(paths[0]!.path).toHaveLength(512);expect(paths[1]!.path).toHaveLength(257)
 chart.chart.literalRadar=radar;chart.chart.literalPie={profile:'literal-pie-v1',firstSliceAngle:0,values:[1],colors:['#000000']};expect(validateNativePptx(deck).ok).toBe(false);delete chart.chart.literalPie
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalRadarPreview:'yes' as never})).rejects.toMatchObject({code:'render.invalidContract'})
 chart.chart.literalRadar={...radar,series:[{...radar.series[0]!,order:0,values:['11','0','0','0']}]};expect(validateNativePptx(deck).ok).toBe(false)
})
