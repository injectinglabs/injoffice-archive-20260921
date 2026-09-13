import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it} from 'vitest'
import {validateNativePptx,type NativePptxDeck,type NativeLiteralBar} from '@injoffice/pptx-native'
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
 chart.chart.literalBar={profile:'literal-bar-v1',barDirection:'column',grouping:'clustered',dataOrigin:'literal',gapWidth:150,overlap:0,categories:['A','B'],series:[{index:1,order:0,values:['-1.25','2e0'],colors:['#FF0000','#0000FF']}],categoryAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},valueAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'-2',max:'3',crossesAt:'0'}}
 return {deck,chart,bar:chart.chart.literalBar}
}
it('compiles and paints literal bars only with their explicit opt-in while preserving source',async()=>{
 const {deck}=fixture(),before=JSON.stringify(deck)
 expect((await compileNativePptxSlide(deck,0,{textLayout:layout})).nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalBarPreview:true})
 expect(tree.nodes[0]!.kind).toBe('group')
 expect(tree.diagnostics.some(d=>d.code==='chart.literalBarPreview')).toBe(true)
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 expect(surface.finish().filter(c=>c.kind==='path')).toHaveLength(2)
 expect(JSON.stringify(deck)).toBe(before)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalBarPreview:true,maxNodes:2})).rejects.toThrow()
})
it('retains packaged fallback when bars cannot fit integer space',async()=>{
 const {deck,chart}=fixture();chart.transform.cx=1
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalBarPreview:true})
 expect(tree.nodes[0]!.kind).toBe('image');expect(tree.diagnostics.some(d=>d.code==='chart.barFrameTooSmall')).toBe(true)
})
it('rejects mismatched source records, competing families and invalid options',async()=>{
 for(const mutate of [(b:NativeLiteralBar)=>{b.series[0]!.values=['1']},(b:NativeLiteralBar)=>{b.valueAxis.min='-Infinity'},(b:NativeLiteralBar)=>{b.categoryAxis.crossAxisId=99},(b:NativeLiteralBar)=>{b.series[0]!.values[0]='1e101'},(b:NativeLiteralBar)=>{b.valueAxis.max='-2'}]){const {deck,bar}=fixture();mutate(bar);expect(validateNativePptx(deck).ok).toBe(false)}
 const {deck,chart}=fixture();chart.chart.literalPie={profile:'literal-pie-v1',firstSliceAngle:0,values:[1],colors:['#000000']};expect(validateNativePptx(deck).ok).toBe(false);delete chart.chart.literalPie
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalBarPreview:'yes' as unknown as boolean})).rejects.toMatchObject({code:'render.invalidContract'})
})
it('bounds the full 16-series by256-category chart within existing node and path budgets',async()=>{
 const {deck,bar}=fixture();bar.categories=Array.from({length:256},(_,i)=>String(i));bar.series=Array.from({length:16},(_,i)=>({index:i,order:i,values:bar.categories.map(()=>'1'),colors:bar.categories.map(()=>'#123456')}))
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalBarPreview:true})
 const group=tree.nodes[0]!;expect(group.kind).toBe('group');if(group.kind==='group')expect(group.children).toHaveLength(4096)
})
