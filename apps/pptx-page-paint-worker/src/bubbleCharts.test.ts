import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview,previewStroke} from './compile.js'
import {validateNativePptx} from '@injoffice/pptx-native'
import {decodePptxPreview} from './contract.js'
const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'pptx-axis-worker-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),digest=`sha256:${createHash('sha256').update(readFileSync(font)).digest('hex')}`
const manifest=resolve(scratch,'fonts.json')
writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:digest}]}))
function fixture(){
 const deck=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8'))
 const chart=deck.slides[0].elements.find((e:{kind:string})=>e.kind==='chart')
 deck.slides=[deck.slides[0]];deck.slides[0].elements=[chart];deck.assets=[];delete chart.chart.previewAssetId
 chart.transform={x:0,y:0,cx:6000000,cy:4000000}
 const style={fontFamily:'DejaVu Sans',fontSize:1200,color:'123456',bold:false,italic:false,language:'en-US'}
 chart.chart.literalBubble={profile:'literal-bubble-v1',dataOrigin:'literal',bubbleScale:100,sizeRepresents:'area',series:[{index:0,order:0,xValues:['-1','2','0'],values:['-1','2','0'],sizes:['1','4','0'],colors:['#CC3300','#3399CC','#000000']}],xAxis:{id:10,crossAxisId:20,position:'b',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,min:'-1',max:'2',crossesAt:'0',labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',style,majorUnit:'1',numberFormat:'0.0'}},yAxis:{id:20,crossAxisId:10,position:'l',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,min:'-1',max:'2',crossesAt:'0',labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',style,majorUnit:'1',numberFormat:'0.0'}}}

 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}
it('uses actual supplied-font labels with fully clipped bubble paths and strict opt-in',async()=>{
 const request=fixture(),before=JSON.stringify(request)
 const off=await compilePptxPreview(request);expect(JSON.stringify(off.nodes)).not.toContain('CC3300')
 const on=await compilePptxPreview({...request,source_chart_preview:true});decodePptxPreview(on)
 expect(on.diagnostics.join(' ')).toContain('chart.bubblePreview');expect(on.diagnostics.join(' ')).toContain('chart.axisLabelsPreview')
 expect(JSON.stringify(on.nodes)).toContain('contentRun');expect(JSON.stringify(on.nodes)).toContain('CC3300');expect(on.font_digests).toEqual([digest]);expect(JSON.stringify(request)).toBe(before)
 request.deck.slides[0].elements[0].chart.literalBubble.xAxis.labels.style.fontFamily='Unavailable'
 const refused=await compilePptxPreview({...request,source_chart_preview:true});expect(refused.diagnostics.join(' ')).toContain('chart.axisLabelsUnavailable');expect(JSON.stringify(refused.nodes)).not.toContain('CC3300')
})
