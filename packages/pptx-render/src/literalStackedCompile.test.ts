import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it} from 'vitest'
import {validateNativePptx,type NativePptxDeck,type NativeLiteralStackedLine} from '@injoffice/pptx-native'
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
 chart.chart.literalStackedLine={profile:'literal-stacked-line-v1',grouping:'stacked',dataOrigin:'literal',categories:['A','B','C'],series:[{index:5,order:0,values:['0','2','0'],color:'#123456',widthEmu:12700}],xAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},yAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'0',max:'1',crossesAt:'0'}}
 return {deck,chart,connected:chart.chart.literalStackedLine}
}

it('requires separate opt-in and retains source while clipping stacked line ink',async()=>{
 const {deck,chart,connected}=fixture();connected.series=[{index:9,order:1,values:['0','2','0'],color:'#123456',widthEmu:12700},{index:7,order:0,values:['0','-1','0'],color:'#ABCDEF',widthEmu:12700}]
 const before=JSON.stringify(deck)
 expect(validateNativePptx(deck).ok).toBe(true)
 expect((await compileNativePptxSlide(deck,0,{textLayout:layout,literalConnectedPreview:true})).nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalStackedPreview:true}),group=tree.nodes[0]!
 if(group.kind!=='group')throw Error('missing group')
 expect(group.clip).toEqual({kind:'rect',rect:{x:0,y:0,cx:chart.transform.cx,cy:chart.transform.cy}})
 expect(group.children).toHaveLength(2)
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const paths=surface.finish().filter(c=>c.kind==='path');expect(paths.map(p=>p.stroke?.color)).toEqual(['#123456','#ABCDEF'])
 expect(paths[0]!.path.filter(c=>c.kind==='moveTo')).toHaveLength(2)
 expect(JSON.stringify(deck)).toBe(before)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,literalStackedPreview:'yes' as unknown as boolean})).rejects.toMatchObject({code:'render.invalidContract'})
})
it('filters singleton and fully clipped stacked lines without invalid path nodes',async()=>{
 for(const singleton of [true,false]){const {deck,connected}=fixture();connected.categories=singleton?['A']:['A','B'];connected.series[0]!.values=singleton?['0']:['2','3'];expect((await compileNativePptxSlide(deck,0,{textLayout:layout,literalStackedPreview:true})).nodes[0]).toMatchObject({kind:'group',children:[]})}
})
it('retains source series order and exact signed percentage bands in bar compiler',async()=>{
 const {deck,chart,connected}=fixture();delete chart.chart.literalStackedLine
 chart.chart.literalStackedBar={profile:'literal-stacked-bar-v1',dataOrigin:'literal',grouping:'percentStacked',overlap:100,gapWidth:150,barDirection:'column',categories:['A'],categoryAxis:connected.xAxis,valueAxis:{...connected.yAxis,min:'-1',max:'1'},series:[{index:9,order:1,values:['-2'],colors:['#123456']},{index:7,order:0,values:['4'],colors:['#ABCDEF']}]}
 const before=JSON.stringify(deck);expect(validateNativePptx(deck).ok).toBe(true)
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout,literalStackedPreview:true}),group=tree.nodes[0]!
 if(group.kind!=='group')throw Error('missing group');expect(group.children).toHaveLength(2)
 expect(group.children.map(p=>p.kind==='shape'?p.fill?.color:undefined)).toEqual(['#123456','#ABCDEF'])
 expect(JSON.stringify(deck)).toBe(before)
 chart.chart.literalConnected=connected as never;expect(validateNativePptx(deck).ok).toBe(false)
})
