import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview} from './compile.js'
import type {PreviewNode} from './contract.js'
const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'pptx-frame-worker-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),digest=`sha256:${createHash('sha256').update(readFileSync(font)).digest('hex')}`,manifest=resolve(scratch,'fonts.json')
writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:digest}]}))
function fixture(grouped=false){
 const deck=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')),table=deck.slides[0].elements.find((e:any)=>e.kind==='table'),group=deck.slides[0].elements.find((e:any)=>e.kind==='group')
 table.graphicFrameLayout='source-anchored-v1';table.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'test.source-frame',message:'source-only frame fixture'}]};table.transform={x:0,y:0,cx:4000000,cy:2000000}
 table.table={columnWidths:[2000000,2000000],rowHeights:[2000000],rows:[['ABC','fgj'].map(text=>({text,fill:'EEEEEE',paragraphs:[{align:'left',level:0,bullet:false,runs:[{text,fontFamily:'DejaVu Sans',fontSizeHundredthPt:1200,color:'123456'}]}],textBody:{leftInsetEmu:10000,rightInsetEmu:10000,topInsetEmu:5000,bottomInsetEmu:5000,wrap:'none',verticalAnchor:'top',autoFit:'none',horizontalOverflow:'clip',verticalOverflow:'overflow'}}))]}
 if(grouped){group.transform={x:1000000,y:1500000,cx:6000000,cy:3000000,rotationAngle:1800000,flipH:true};group.childTransform={x:0,y:0,cx:4000000,cy:2000000};group.children=[table];group.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'test.source-frame',message:'source-only frame fixture'}]}}
 deck.slides=[deck.slides[0]];deck.slides[0].elements=[grouped?group:table];deck.assets=[];deck.compatibility.status='preserveOnly';deck.slides[0].compatibility.status='preserveOnly'
 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}
type Matrix=readonly number[]
function multiply(a:Matrix,b:Matrix):number[]{return [a[0]!*b[0]!+a[2]!*b[1]!,a[1]!*b[0]!+a[3]!*b[1]!,a[0]!*b[2]!+a[2]!*b[3]!,a[1]!*b[2]!+a[3]!*b[3]!,a[0]!*b[4]!+a[2]!*b[5]!+a[4]!,a[1]!*b[4]!+a[3]!*b[5]!+a[5]!]}
function glyphs(nodes:readonly PreviewNode[],parent:Matrix=[1,0,0,1,0,0]):number[][]{return nodes.flatMap(node=>node.kind==='group'?glyphs(node.children,multiply(parent,node.transform)):node.kind==='path'&&node.fill==='123456'?[[...parent]]:[])}
it('paints intrinsic table tracks and unscaled supplied glyphs under nonuniform reflected rotated source groups',async()=>{
 const baseline=await compilePptxPreview(fixture()),input=fixture(true),before=JSON.stringify(input),result=await compilePptxPreview(input)
 const a=glyphs(baseline.nodes),b=glyphs(result.nodes);expect(a).toHaveLength(6);expect(b,JSON.stringify(result)).toHaveLength(6)
 for(let i=0;i<6;i++)expect(b[i]!.slice(0,4)).toEqual(a[i]!.slice(0,4))
 expect(JSON.stringify(result.nodes)).toContain('clip');expect(result.diagnostics.join(' ')).toContain('graphicFrame.sourceAnchoredPreview');expect(JSON.stringify(input)).toBe(before)
})
it('retains source frame and intrinsic grid mismatch without resizing glyphs or tracks',async()=>{
 const input=fixture();input.deck.slides[0].elements[0].transform.cx=6000000
 const result=await compilePptxPreview(input);expect(glyphs(result.nodes)).toHaveLength(6)
 expect(input.deck.slides[0].elements[0].table.columnWidths).toEqual([2000000,2000000])
})
function nestedFixture(depth:number){
 const input=fixture(true),prototype=input.deck.slides[0].elements[0],table=prototype.children[0]
 let child=table
 for(let i=0;i<depth;i++){
  const group=structuredClone(prototype)
  group.id=`nested-${i}`;group.source.objectId=String(100+i)
  group.transform={x:0,y:0,cx:4000000,cy:2000000};group.childTransform={x:0,y:0,cx:4000000,cy:2000000};group.children=[child];child=group
 }
 input.deck.slides[0].elements=[child];return input
}
it('qualifies complete glyph and finite clip spans at the deepest paintable source hierarchy',async()=>{
 const result=await compilePptxPreview(nestedFixture(25))
 expect(glyphs(result.nodes)).toHaveLength(6)
 expect(JSON.stringify(result.nodes)).toContain('clip')
 await expect(compilePptxPreview(nestedFixture(26))).rejects.toThrow('contract nesting exceeds 64')
})
function chartFixture(){
 const deck=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8'))
 const chart=deck.slides[0].elements.find((e:{kind:string})=>e.kind==='chart')
 deck.slides=[deck.slides[0]];deck.slides[0].elements=[chart];deck.assets=[];delete chart.chart.previewAssetId
 chart.transform={x:0,y:0,cx:6000000,cy:4000000}
 const style={fontFamily:'DejaVu Sans',fontSize:1200,color:'123456',bold:false,italic:false,language:'en-US'}
 chart.chart.literalBubble={profile:'literal-bubble-v1',dataOrigin:'literal',bubbleScale:100,sizeRepresents:'area',series:[{index:0,order:0,xValues:['-1','2','0'],values:['-1','2','0'],sizes:['1','4','0'],colors:['#CC3300','#3399CC','#000000']}],xAxis:{id:10,crossAxisId:20,position:'b',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,min:'-1',max:'2',crossesAt:'0',labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',style,majorUnit:'1',numberFormat:'0.0'}},yAxis:{id:20,crossAxisId:10,position:'l',orientation:'minMax',deleted:false,color:'#000000',widthEmu:12700,min:'-1',max:'2',crossesAt:'0',labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',style,majorUnit:'1',numberFormat:'0.0'}}}

 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}

it('preserves actual axis glyph metrics for fractional physical chart layout and bounds all 4096 circles',async()=>{
 const request=chartFixture(),chart=request.deck.slides[0].elements[0]
 chart.graphicFrameLayout='source-anchored-v1'
 const baseline=await compilePptxPreview({...request,source_chart_preview:true})
 expect(glyphs(baseline.nodes).length).toBeGreaterThan(10)
 const group=fixture(true).deck.slides[0].elements[0]
 group.transform={x:1000000,y:1000000,cx:6000001,cy:4000001,rotationAngle:1800000,flipV:true}
 group.childTransform={x:0,y:0,cx:12000000,cy:8000000}
 group.children=[chart];request.deck.slides[0].elements=[group]
 const before=JSON.stringify(request),result=await compilePptxPreview({...request,source_chart_preview:true})
 expect(result.diagnostics.join(' ')).toContain('nearest-EMU physical dimensions')
 expect(glyphs(result.nodes).length).toBeGreaterThan(10)
 expect(glyphs(result.nodes)[0]!.slice(0,4)).toEqual(glyphs(baseline.nodes)[0]!.slice(0,4))
 expect(JSON.stringify(request)).toBe(before)
 const series=chart.chart.literalBubble.series[0]
 series.xValues=Array(256).fill('0');series.values=Array(256).fill('0');series.sizes=Array(256).fill('1');series.colors=Array(256).fill('#CC3300')
 chart.chart.literalBubble.series=Array.from({length:16},(_,index)=>({...series,index,order:index}))
 const large=await compilePptxPreview({...request,source_chart_preview:true})
 const circles=(nodes:readonly PreviewNode[]):number=>nodes.reduce((n,node)=>n+(node.kind==='group'?circles(node.children):node.kind==='path'&&node.fill==='CC3300'?1:0),0)
 expect(circles(large.nodes)).toBe(4096)
})
it('rejects untrusted profile authority, missing profile mismatch, and unknown profile spellings',async()=>{
 for(const change of [
  (e:any)=>{e.provenance='authored'},
  (e:any)=>{delete e.source},
  (e:any)=>{e.compatibility={status:'editable',diagnostics:[]}},
  (e:any)=>{e.graphicFrameLayout='source-anchored-v2'},
  (e:any)=>{delete e.graphicFrameLayout;e.transform.cx=5000000},
 ]){
  const input=fixture();change(input.deck.slides[0].elements[0])
  await expect(compilePptxPreview(input)).rejects.toThrow('invalid native PPTX contract')
 }
})
it('retains an unavailable-glyph refusal placeholder inside the deepest qualified table clip',async()=>{
 const input=nestedFixture(25)
 let leaf=input.deck.slides[0].elements[0];while(leaf.children?.length)leaf=leaf.children[0]
 leaf.table.rows[0][0].text=leaf.table.rows[0][0].paragraphs[0].runs[0].text='\u{10ffff}'
 const result=await compilePptxPreview(input)
 expect(glyphs(result.nodes)).toHaveLength(3)
 expect(result.diagnostics.join(' ')).toContain('text.refused')
 expect(JSON.stringify(result.nodes)).toContain('clip')
})
