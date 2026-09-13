import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it,vi} from 'vitest'
import {validateNativePptx,type NativePptxDeck} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,type NativePptxTextLayout} from './index.js'
import * as geometry from './geometry.js'
const layout:NativePptxTextLayout={
 manifest:{version:1,manifestId:'geometry-only',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused',contentDigest:`sha256:${'0'.repeat(64)}`}}],fallbackChains:[]},
 resolver:{providerId:'unused',providerRevision:'1',resolve(){throw Error('No text')},load(){throw Error('No font')}},shaper:{providerId:'unused',providerRevision:'1',shape(){throw Error('No glyphs')}},defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'},
}
function fixture(values=[1,3]){
 const deck=JSON.parse(readFileSync(resolve(import.meta.dirname,'../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')) as NativePptxDeck
 const chart=deck.slides[0]!.elements.find(e=>e.kind==='chart')!
 if(chart.kind!=='chart')throw Error('missing chart')
 deck.slides[0]!.elements=[chart]
 chart.chart.literalDoughnut={profile:'literal-doughnut-v1',firstSliceAngle:90,holeSize:50,values,colors:values.map(()=> '#0000FF')}
 return {deck,chart}
}
it.each([[1],[999,1],Array(64).fill(1)].map(values=>({values})))('compiles a complete annulus at actual path/node budgets ($values)',async({values})=>{
 const {deck}=fixture(values),before=JSON.stringify(deck)
 const normal=await compileNativePptxSlide(deck,0,{textLayout:layout})
 expect(normal.nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalDoughnutPreview:true})
 expect(tree.nodes[0]!.kind).toBe('group')
 expect(tree.diagnostics.some(d=>d.code==='chart.literalDoughnutPreview')).toBe(true)
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const paths=surface.finish().filter(c=>c.kind==='path')
 expect(paths).toHaveLength(values.length)
 expect(paths.every(p=>p.path.length<=363)).toBe(true)
 expect(paths.reduce((n,p)=>n+p.path.length,0)).toBeLessThanOrEqual(680)
 expect(JSON.stringify(deck)).toBe(before)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalDoughnutPreview:true,maxNodes:values.length})).rejects.toThrow()
})
it('retains fallback for an unrepresentable small hole and requires its own opt-in',async()=>{
 const {deck,chart}=fixture()
 const pieOnly=await compileNativePptxSlide(deck,0,{textLayout:layout,literalPiePreview:true})
 expect(pieOnly.nodes[0]!.kind).toBe('image')
 chart.transform.cx=1;chart.transform.cy=1
 const tiny=await compileNativePptxSlide(deck,0,{textLayout:layout,literalDoughnutPreview:true})
 expect(tiny.nodes[0]!.kind).toBe('image')
 expect(tiny.diagnostics.some(d=>d.code==='chart.doughnutFrameTooSmall')).toBe(true)
})
it('rejects mismatched values, competing families and invalid public options',async()=>{
 const {deck,chart}=fixture()
 chart.chart.literalDoughnut!.colors=[]
 expect(validateNativePptx(deck).ok).toBe(false)
 chart.chart.literalDoughnut!.colors=['#0000FF','#0000FF']
 chart.chart.literalPie={profile:'literal-pie-v1',firstSliceAngle:0,values:[1],colors:['#000000']}
 expect(validateNativePptx(deck).ok).toBe(false)
 delete chart.chart.literalPie
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalDoughnutPreview:'yes' as unknown as boolean})).rejects.toMatchObject({code:'render.invalidContract'})
})
it('rejects 513 generated path commands at the compiler guard',async()=>{
 const {deck,chart}=fixture()
 // Inject an oversized geometry result to exercise the compiler's actual guard;
 // no source/custom-path contract is added just to manufacture this boundary.
 deck.slides[0]!.elements=[{kind:'shape',id:chart.id,provenance:'authored',preset:'rect',transform:chart.transform,fill:'0000FF',paragraphs:[],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}]
 const spy=vi.spyOn(geometry,'presetPath').mockReturnValue(Array.from({length:513},()=>({kind:'lineTo' as const,x:0,y:0})))
 try{await expect(compileNativePptxSlide(deck,0,{textLayout:layout})).rejects.toMatchObject({code:'render.pathBudget'})}finally{spy.mockRestore()}
})
