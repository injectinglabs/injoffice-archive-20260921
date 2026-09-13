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
 chart.chart.literalConnected={profile:'literal-line-v1',dataOrigin:'literal',categories:['Alpha','Beta','Gamma'],series:[{index:0,order:0,values:['-1','2','0'],color:'#CC3300',widthEmu:12700}],xAxis:{id:10,crossAxisId:20,position:'b',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',style}},yAxis:{id:20,crossAxisId:10,position:'l',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,min:'-1',max:'2',crossesAt:'0',labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'out',style,majorUnit:'1',numberFormat:'0.0'}}}
 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}
it('paints source axis labels from exact operator font outlines only after explicit opt-in',async()=>{
 const request=fixture(),before=JSON.stringify(request)
 const off=await compilePptxPreview(request)
 expect(off.source_chart_preview).toBeUndefined();expect(JSON.stringify(off.nodes)).not.toContain('contentRun')
 const on=await compilePptxPreview({...request,source_chart_preview:true})
 expect(on.source_chart_preview).toBe(true);expect(on.chart_axis_layout_policy).toBe('supplied-outline-margins-v1')
 expect(on.diagnostics.join('\n')).toContain('chart.axisLabelsPreview')
 expect(on.diagnostics.join('\n')).not.toContain('axisLabelsUnavailable')
 expect(JSON.stringify(on.nodes)).toContain('contentRun');expect(JSON.stringify(on.nodes)).toContain('CC3300')
 expect(on.font_digests).toEqual([digest]);expect(JSON.stringify(request)).toBe(before)
 for(const delta of [{source_chart_preview:false},{source_chart_preview:'true'},{source_chart_preview:undefined},{chart_axis_layout_policy:'other'}])expect(()=>decodePptxPreview({...on,...delta})).toThrow()
 await expect(compilePptxPreview({...request,source_chart_preview:'true'})).rejects.toThrow('boolean')
})
it('refuses labels with unavailable fonts or insufficient space without silently painting unlabeled data',async()=>{
 for(const tiny of [false,true]){
  const request=fixture(),chart=request.deck.slides[0].elements[0]
  if(tiny)chart.transform.cx=100000;else chart.chart.literalConnected.xAxis.labels.style.fontFamily='Unavailable'
  const result=await compilePptxPreview({...request,source_chart_preview:true})
  expect(result.diagnostics.join('\n')).toContain('chart.axisLabelsUnavailable')
  expect(JSON.stringify(result.nodes)).not.toContain('CC3300')
 }
})

it('rejects label contract drift, cache-like formats, controls and crossing ambiguity',()=>{
 for(const mutate of [
  (a:any)=>{a.labels.style.language='en-US\n'},
  (a:any)=>{a.labels.style.fontFamily='+mn-lt'},
  (a:any)=>{a.labels.style.color='#123456'},
  (a:any)=>{a.labels.majorUnit='0'},
  (a:any)=>{a.labels.majorUnit='0.00001'},
  (a:any)=>{a.labels.numberFormat='General'},
  (a:any)=>{a.deleted=true},
 ]){const request=fixture();mutate(request.deck.slides[0].elements[0].chart.literalConnected.yAxis);expect(validateNativePptx(request.deck).ok).toBe(false)}
 const request=fixture();request.deck.slides[0].elements[0].chart.literalConnected.xAxis.labels.majorTickMark='out';expect(validateNativePptx(request.deck).ok).toBe(false)
})
it('refuses unsupported axis scripts without painting arbitrary Latin-shaper output',async()=>{
 const request=fixture();request.deck.slides[0].elements[0].chart.literalConnected.categories[0]='שלום'
 const result=await compilePptxPreview({...request,source_chart_preview:true})
 expect(result.diagnostics.join(' ')).toContain('supported Latin/common')
 expect(JSON.stringify(result.nodes)).not.toContain('CC3300')
})

it('normalizes only qualified chart RGB at the paint transport boundary',async()=>{
 expect(previewStroke({color:'#123456',widthEmu:1}).stroke).toBe('123456')
 expect(previewStroke({color:'123456',widthEmu:1}).stroke).toBe('123456')
 expect(previewStroke({color:'#bad',widthEmu:1}).stroke).toBe('#bad')
 const request=fixture(),chart=request.deck.slides[0].elements[0],c=chart.chart.literalConnected
 chart.chart.literalBar={profile:'literal-bar-v1',barDirection:'column',grouping:'clustered',dataOrigin:'literal',gapWidth:150,overlap:0,categories:c.categories,categoryAxis:c.xAxis,valueAxis:c.yAxis,series:[{index:0,order:0,values:['-1','2','0'],colors:['#AA0000','#00BB00','#0000CC']}]};delete chart.chart.literalConnected
 const result=await compilePptxPreview({...request,source_chart_preview:true}),raw=JSON.stringify(result.nodes)
 expect(raw).toContain('"fill":"AA0000"');expect(raw).toContain('"fill":"00BB00"');expect(raw).not.toContain('#')
 expect(result.diagnostics.join(' ')).toContain('chart.axisLabelsPreview')
})
