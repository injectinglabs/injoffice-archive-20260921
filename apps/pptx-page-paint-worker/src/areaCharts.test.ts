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
 chart.chart.literalArea={profile:'literal-area-v1',dataOrigin:'literal',grouping:'standard',categories:['Alpha','Beta','Gamma'],series:[{index:0,order:0,values:['-1','2','0'],color:'#CC3300'}],xAxis:{id:10,crossAxisId:20,position:'b',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',style}},yAxis:{id:20,crossAxisId:10,position:'l',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,min:'-1',max:'2',crossesAt:'0',labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'out',style,majorUnit:'1',numberFormat:'0.0'}}}
 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}
it('paints area source values and supplied-font labels only after literal opt-in',async()=>{
 const request=fixture(),before=JSON.stringify(request)
 const off=await compilePptxPreview(request);expect(JSON.stringify(off.nodes)).not.toContain('CC3300')
 const on=await compilePptxPreview({...request,source_chart_preview:true});decodePptxPreview(on)
 expect(on.diagnostics.join(' ')).toContain('chart.literalAreaPreview');expect(on.diagnostics.join(' ')).toContain('chart.axisLabelsPreview')
 expect(JSON.stringify(on.nodes)).toContain('contentRun');expect(JSON.stringify(on.nodes)).toContain('CC3300');expect(on.font_digests).toEqual([digest]);expect(JSON.stringify(request)).toBe(before)
})
it('replays all1280 compound commands through bounded SVG transport without a hidden512 cap',async()=>{
 const request=fixture(),chart=request.deck.slides[0].elements[0],a=chart.chart.literalArea,values=Array.from({length:256},(_,i)=>i%2?'1':'-1')
 delete a.xAxis.labels;delete a.yAxis.labels;a.categories=values;a.series[0].values=values;a.yAxis.min='-.5';a.yAxis.max='.5'
 const on=await compilePptxPreview({...request,source_chart_preview:true});decodePptxPreview(on)
 const paths:any[]=[];const visit=(nodes:any[])=>{for(const node of nodes){if(node.kind==='path')paths.push(node);if(node.kind==='group')visit(node.children)}};visit(on.nodes)
 const fills=paths.filter(p=>p.fill==='CC3300');expect(fills).toHaveLength(1)
 expect(fills[0].d.match(/[MLZ]/g)).toHaveLength(1280);expect(fills[0].d.length).toBeLessThan(200000)
 a.grouping='percentStacked';a.series[0].values=values.map(()=>'0')
 const zero=await compilePptxPreview({...request,source_chart_preview:true});decodePptxPreview(zero);expect(JSON.stringify(zero.nodes)).not.toContain('CC3300');expect(JSON.stringify(zero.nodes)).toContain('000000')
})
it('retains labeled area fallback on missing fonts or insufficient plot space',async()=>{
 for(const tiny of [false,true]){
  const request=fixture(),chart=request.deck.slides[0].elements[0]
  if(tiny)chart.transform.cx=100000;else chart.chart.literalArea.xAxis.labels.style.fontFamily='Unavailable'
  const result=await compilePptxPreview({...request,source_chart_preview:true})
  expect(result.diagnostics.join(' ')).toContain('chart.axisLabelsUnavailable');expect(JSON.stringify(result.nodes)).not.toContain('CC3300')
 }
})
