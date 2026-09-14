import {evaluatedGeometryPaths} from './evaluatedGeometry.js'
import {PPTX_RENDER_LIMITS} from './types.js'
import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it} from 'vitest'
import {validateNativePptx,type NativePptxDeck,type NativeLiteralArea} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,type NativePptxTextLayout} from './index.js'
const layout:NativePptxTextLayout={
 manifest:{version:1,manifestId:'geometry-only',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused',contentDigest:`sha256:${'0'.repeat(64)}`}}],fallbackChains:[]},
 resolver:{providerId:'unused',providerRevision:'1',resolve(){throw Error('No text')},load(){throw Error('No font')}},shaper:{providerId:'unused',providerRevision:'1',shape(){throw Error('No glyphs')}},defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'},
}

function fixture(){
 const deck=JSON.parse(readFileSync(resolve(import.meta.dirname,'../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')) as NativePptxDeck
 const chart=deck.slides[0]!.elements.find(e=>e.kind==='chart')!
 if(chart.kind!=='chart')throw Error('missing chart')
 deck.slides[0]!.elements=[chart]
 const area:NativeLiteralArea={profile:'literal-area-v1',dataOrigin:'literal',grouping:'standard',categories:['A','B','C'],series:[{index:3,order:0,color:'#123456',values:['-1','1','-1']}],xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:false,color:'#000000',widthEmu:12700},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:false,color:'#000000',widthEmu:12700,min:'-2',max:'2',crossesAt:'0'}}
 chart.chart.literalArea=area
 return {deck,chart,area}
}
it('compiles complete same-series compound fills only with area opt-in and retains source',async()=>{
 const {deck}=fixture(),before=JSON.stringify(deck)
 expect((await compileNativePptxSlide(deck,0,{textLayout:layout})).nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalAreaPreview:true})
 expect(tree.nodes[0]!.kind).toBe('group');expect(tree.diagnostics.some(d=>d.code==='chart.literalAreaPreview')).toBe(true)
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const paths=surface.finish().filter(c=>c.kind==='path');expect(paths).toHaveLength(3)
 expect(paths[0]!.path.filter(c=>c.kind==='moveTo')).toHaveLength(3)
 expect(paths[0]!.fill).toBe('#123456');expect(paths[0]!.stroke).toBeUndefined()
 expect(JSON.stringify(deck)).toBe(before)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalAreaPreview:true,maxNodes:2})).rejects.toThrow()
})
it('retains axes for zero percent totals and authored order for overlapping standard fills',async()=>{
 const {deck,chart,area}=fixture()
 chart.chart.literalArea={...area,grouping:'percentStacked',series:[{...area.series[0]!,values:['0','-0','0.0']}]}
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalAreaPreview:true}),surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const paths=surface.finish().filter(c=>c.kind==='path');expect(paths).toHaveLength(2);expect(paths.every(p=>!p.fill)).toBe(true)
 chart.chart.literalArea={...area,series:[area.series[0]!,{...area.series[0]!,index:9,order:1,color:'#ABCDEF'}]}
 const next=await compileNativePptxSlide(deck,0,{textLayout:layout,literalAreaPreview:true}),paint=createRecordingPaintSurface();paintSlideRenderTree(next,paint)
 expect(paint.finish().filter(c=>c.kind==='path'&&c.fill).map(c=>c.kind==='path'?c.fill:'')).toEqual(['#123456','#ABCDEF'])
})
it('admits the qualified1280-command path without expanding arbitrary source geometry',async()=>{
 const {deck,chart,area}=fixture(),values=Array.from({length:256},(_,i)=>i%2?'1':'-1')
 chart.chart.literalArea={...area,categories:values,series:[{...area.series[0]!,values}],yAxis:{...area.yAxis,min:'-.5',max:'.5'}}
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalAreaPreview:true}),surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 expect(surface.finish().find(c=>c.kind==='path'&&c.fill)?.kind).toBe('path')
 const fill=surface.finish().find(c=>c.kind==='path'&&c.fill);if(fill?.kind==='path')expect(fill.path).toHaveLength(1280)
 expect(PPTX_RENDER_LIMITS.maxPathCommands).toBe(512)
 const invalid={profile:'drawingml-paths-v1',textRect:{x:0,y:0,cx:1,cy:1},paths:[{fillMode:'norm',stroke:false,commands:[{kind:'moveTo',x:0,y:0},...Array(512).fill({kind:'close'})]}]}
 expect(()=>evaluatedGeometryPaths(invalid as never,()=>{},'$.geometry',()=>{})).toThrow('geometry path budget exceeded')
})
it('rejects foreign provenance, competing families and nonboolean opt-in',async()=>{
 const {deck,chart,area}=fixture()
 for(const delta of [{grouping:'stacked'},{dataOrigin:'embedded-workbook'},{series:[{...area.series[0]!,order:-0}]}]){chart.chart.literalArea={...area,...delta} as NativeLiteralArea;expect(validateNativePptx(deck).ok).toBe(false)}
 chart.chart.literalArea=area;chart.chart.literalPie={profile:'literal-pie-v1',firstSliceAngle:0,values:[1],colors:['#000000']};expect(validateNativePptx(deck).ok).toBe(false);delete chart.chart.literalPie
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalAreaPreview:'yes' as never})).rejects.toMatchObject({code:'render.invalidContract'})
})
