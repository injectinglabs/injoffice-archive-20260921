import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it} from 'vitest'
import {validateNativePptx,type NativePptxDeck,type NativeLiteralConnected} from '@injoffice/pptx-native'
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
 chart.chart.literalConnected={profile:'literal-line-v1',dataOrigin:'literal',categories:['A','B','C'],series:[{index:5,order:0,values:['0','2','0'],color:'#123456',widthEmu:12700}],xAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},yAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'0',max:'1',crossesAt:'0'}}
 return {deck,chart,connected:chart.chart.literalConnected}
}
it('requires opt-in, retains source, and clips stroke envelopes at the plot frame',async()=>{
 const {deck,chart}=fixture(),before=JSON.stringify(deck)
 expect((await compileNativePptxSlide(deck,0,{textLayout:layout})).nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalConnectedPreview:true})
 const group=tree.nodes[0]!;expect(group.kind).toBe('group');if(group.kind!=='group')throw Error('missing group')
 expect(group.clip).toEqual({kind:'rect',rect:{x:0,y:0,cx:chart.transform.cx,cy:chart.transform.cy}})
 expect(group.children).toHaveLength(1)
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const paths=surface.finish().filter(c=>c.kind==='path');expect(paths).toHaveLength(1);expect(paths[0]!.path.filter(c=>c.kind==='moveTo')).toHaveLength(2)
 expect(paths[0]!.stroke).toMatchObject({color:'#123456',cap:'flat',join:'round'})
 expect(JSON.stringify(deck)).toBe(before)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalConnectedPreview:true,maxNodes:1})).rejects.toThrow()
})
it('compiles singleton and fully clipped series as empty groups without invalid shapes',async()=>{
 for(const singleton of [false,true]){
  const {deck,connected}=fixture();connected.categories=singleton?['A']:['A','B'];connected.series[0]!.values=singleton?['0']:['2','3']
  const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalConnectedPreview:true}),surface=createRecordingPaintSurface()
  expect(tree.nodes[0]).toMatchObject({kind:'group',children:[]});paintSlideRenderTree(tree,surface);expect(surface.finish().filter(c=>c.kind==='path')).toHaveLength(0)
 }
})
it('compiles the full source budget and510-command clipped paths',async()=>{
 const {deck,connected}=fixture();connected.categories=Array.from({length:256},(_,i)=>String(i));connected.series=Array.from({length:16},(_,i)=>({index:i,order:i,values:connected.categories.map((_,j)=>j%2?'2':'-1'),color:'#123456',widthEmu:12700}))
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalConnectedPreview:true}),node=tree.nodes[0]!
 if(node.kind!=='group')throw Error('missing group');expect(node.children).toHaveLength(16)
 for(const child of node.children){if(child.kind!=='shape')throw Error('missing shape');expect(child.path).toHaveLength(510)}
})
it('rejects family invariant drift and competing profiles across the public validator',async()=>{
 for(const mutate of [(c:NativeLiteralConnected)=>{c.series[0]!.xValues=[]},(c:NativeLiteralConnected)=>{c.categories=[]},(c:NativeLiteralConnected)=>{c.profile='literal-scatter-v1'},(c:NativeLiteralConnected)=>{c.series[0]!.values=['1']},(c:NativeLiteralConnected)=>{c.yAxis.crossesAt='1'},(c:NativeLiteralConnected)=>{c.series[0]!.values[0]='1e101'},(c:NativeLiteralConnected)=>{c.xAxis.min='-1'}]){const {deck,connected}=fixture();mutate(connected);expect(validateNativePptx(deck).ok).toBe(false)}
 const {deck,chart}=fixture();chart.chart.literalPie={profile:'literal-pie-v1',firstSliceAngle:0,values:[1],colors:['#000000']};expect(validateNativePptx(deck).ok).toBe(false);delete chart.chart.literalPie
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalConnectedPreview:'yes' as unknown as boolean})).rejects.toMatchObject({code:'render.invalidContract'})
})
it('compiles explicit XY values with both scale reversals and no category substitution',async()=>{
 const {deck,connected}=fixture();connected.profile='literal-scatter-v1';connected.categories=[];connected.series[0]!.xValues=['1','-1','1'];connected.series[0]!.values=['-1','0','1'];Object.assign(connected.xAxis,{min:'-1',max:'1',crossesAt:'0',orientation:'maxMin'});Object.assign(connected.yAxis,{min:'-1',max:'1',orientation:'maxMin'})
 expect(validateNativePptx(deck).ok).toBe(true)
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalConnectedPreview:true}),group=tree.nodes[0]!
 if(group.kind!=='group'||group.children[0]?.kind!=='shape')throw Error('missing path')
 expect(group.children[0].path[0]).toEqual({kind:'moveTo',x:0,y:0})
})
