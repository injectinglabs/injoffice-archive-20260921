import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview} from './compile.js'
import type {PreviewNode} from './contract.js'
const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'pptx-affine-worker-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),digest=`sha256:${createHash('sha256').update(readFileSync(font)).digest('hex')}`,manifest=resolve(scratch,'fonts.json')
writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:digest}]}))
function fixture(flipH=false,flipV=false,rotationAngle=0){
 const deck=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')),shape=deck.slides[0].elements.find((e:{kind:string})=>e.kind==='shape')
 deck.slides=[deck.slides[0]];deck.slides[0].elements=[shape];deck.assets=[]
 shape.transform={x:1000000,y:2000000,cx:4000000,cy:2000000,rotationAngle,flipH,flipV};shape.preset='triangle';shape.fill='336699';delete shape.stroke
 shape.paragraphs=[{align:'left',level:0,bullet:false,runs:[{text:'ABCDEF',fontFamily:'DejaVu Sans',fontSizeHundredthPt:1200,color:'123456'}]}]
 shape.textBody={leftInsetEmu:100000,rightInsetEmu:200000,topInsetEmu:30000,bottomInsetEmu:50000,wrap:'none',verticalAnchor:'top',autoFit:'none',horizontalOverflow:'overflow',verticalOverflow:'overflow'}
 shape.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'test.preview',message:'Read-only source affine'}]};deck.compatibility.status='preserveOnly';deck.slides[0].compatibility.status='preserveOnly'
 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}
type Matrix=readonly number[]
function multiply(a:Matrix,b:Matrix):number[]{return [a[0]!*b[0]!+a[2]!*b[1]!,a[1]!*b[0]!+a[3]!*b[1]!,a[0]!*b[2]!+a[2]!*b[3]!,a[1]!*b[2]!+a[3]!*b[3]!,a[0]!*b[4]!+a[2]!*b[5]!+a[4]!,a[1]!*b[4]!+a[3]!*b[5]!+a[5]!]}
function paths(nodes:readonly PreviewNode[],parent:Matrix=[1,0,0,1,0,0]):{fill:string;matrix:number[]}[]{return nodes.flatMap(node=>node.kind==='group'?paths(node.children,multiply(parent,node.transform)):node.kind==='path'?[{fill:node.fill??'',matrix:[...parent]}]:[])}
it('retains readable glyph orientation for H, V and combined flips with asymmetric text insets',async()=>{
 const first=await compilePptxPreview(fixture()),baseline=paths(first.nodes).filter(path=>path.fill==='123456')
 expect(baseline,JSON.stringify(first)).toHaveLength(6)
 for(const [h,v] of [[true,false],[false,true],[true,true]]){
  const result=await compilePptxPreview(fixture(h,v)),all=paths(result.nodes),glyphs=all.filter(path=>path.fill==='123456'),outline=all.find(path=>path.fill==='336699')!
  expect(outline.matrix[0]).toBe(h?-1:1);expect(outline.matrix[3]).toBe(v?-1:1)
  expect(glyphs).toHaveLength(6)
  for(let i=0;i<glyphs.length;i++){
   const got=glyphs[i]!.matrix,original=baseline[i]!.matrix
   expect(got[0]).toBeCloseTo(original[0]!*(v?-1:1),10);expect(got[3]).toBeCloseTo(original[3]!*(v?-1:1),10)
   expect(got[4]).toBeCloseTo(v?6000000-original[4]!:original[4]!,7)
   expect(got[5]).toBeCloseTo(v?6000000-original[5]!:original[5]!,7)
  }
 }
})
it('paints supplied outline glyphs under a noncardinal source affine through the worker boundary',async()=>{
 const before=fixture(true,false,1800000),copy=JSON.stringify(before),result=await compilePptxPreview(before),all=paths(result.nodes),glyphs=all.filter(path=>path.fill==='123456')
 expect(glyphs).toHaveLength(6)
 const [a,b]=glyphs[0]!.matrix
 expect(b!/a!).toBeCloseTo(1/Math.sqrt(3),12)
 expect(JSON.stringify(before)).toBe(copy);expect(result.font_digests).toEqual([digest])
})

it('composes the primary +90 shape / -90 body example without rotating supplied glyph metrics',async()=>{
 const baseline=paths((await compilePptxPreview(fixture())).nodes).filter(path=>path.fill==='123456')
 const input=fixture(false,false,5400000)
 input.deck.slides[0].elements[0].textBody.rotationAngle60000=-5400000
 const snapshot=JSON.stringify(input),result=await compilePptxPreview(input)
 const glyphs=paths(result.nodes).filter(path=>path.fill==='123456')
 expect(glyphs,JSON.stringify(result)).toHaveLength(6)
 for(let i=0;i<glyphs.length;i++)for(let coordinate=0;coordinate<6;coordinate++)expect(glyphs[i]!.matrix[coordinate]).toBeCloseTo(baseline[i]!.matrix[coordinate]!,7)
 expect(JSON.stringify(input)).toBe(snapshot)
})
it('upright cancels source rotation and reflection while retaining actual positive font axes',async()=>{
 const baseline=paths((await compilePptxPreview(fixture())).nodes).filter(path=>path.fill==='123456')
 for(const [h,v,angle] of [[false,false,5400000],[true,false,1800000],[false,true,10800000],[true,true,16200000]] as const){
  const input=fixture(h,v,angle);input.deck.slides[0].elements[0].textBody.upright=true
  input.deck.slides[0].elements[0].textBody.rotationAngle60000=1800000
  const result=await compilePptxPreview(input),glyphs=paths(result.nodes).filter(path=>path.fill==='123456')
  expect(glyphs,JSON.stringify(result)).toHaveLength(6)
  for(let i=0;i<glyphs.length;i++)for(let axis=0;axis<4;axis++)expect(glyphs[i]!.matrix[axis]).toBeCloseTo(baseline[i]!.matrix[axis]!,7)
 }
})

it('keeps upright actual glyph scale under an anisotropic parsed group with a fractional area',async()=>{
 const input=fixture(),shape=input.deck.slides[0].elements[0]
 shape.transform={x:0,y:0,cx:4000000,cy:2000000,rotationAngle:5400000}
 shape.textBody.upright=true;shape.textBody.verticalAnchor='center';shape.paragraphs[0].align='center'
 const original=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8'))
 const group=original.slides[0].elements.find((element:{kind:string})=>element.kind==='group')
 group.transform={x:2000000,y:2000000,cx:6000000,cy:2000000};group.childTransform={x:0,y:0,cx:4000000,cy:2000000}
 group.children=[shape];group.compatibility=shape.compatibility;input.deck.slides[0].elements=[group]
 const baseline=paths((await compilePptxPreview(fixture())).nodes).filter(path=>path.fill==='123456')
 const result=await compilePptxPreview(input),glyphs=paths(result.nodes).filter(path=>path.fill==='123456')
 expect(glyphs,JSON.stringify(result)).toHaveLength(6)
 for(let i=0;i<glyphs.length;i++){
  expect(glyphs[i]!.matrix[0]).toBeCloseTo(baseline[i]!.matrix[0]!*1.5,7)
  expect(glyphs[i]!.matrix[3]).toBeCloseTo(baseline[i]!.matrix[3]!,7)
  expect(glyphs[i]!.matrix[1]).toBeCloseTo(0,7);expect(glyphs[i]!.matrix[2]).toBeCloseTo(0,7)
 }
})
