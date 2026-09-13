import {readFileSync} from 'node:fs'
import {createRequire} from 'node:module'
import {describe,it,expect} from 'vitest'
import {compileNativeDocxTextboxShapeV1} from './nativeTextboxShapePaintV1.js'
import {decodeNativeDocxTextboxGeometryV1 as decode,decodeNativeDocxTextboxShapePaintV1 as decodePaint,nativeTextboxFontDigestV1,nativeTextboxGeometryPlainData,type NativeDocxTextboxGeometryEvidenceV1} from './nativeTextboxGeometryPreviewV1.js'
const original=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8'))
const require=createRequire(import.meta.url),font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
function fixture(){
 const document=structuredClone(original),p=document.body.blocks[0].paragraph,anchor={...p.anchor,path:p.anchor.path+'/w:r[1]/w:drawing[1]/wp:inline[1]/a:graphic[1]/a:graphicData[1]',start_byte:200,end_byte:400}
 document.unsupported=[{id:'shape:1',code:'PICTURE_GRAPHIC_REQUIRED',capability:'drawings',scope_id:p.id,anchor,preservation:'refuse-mutation',message:'Drawing preserved'}]
 const evidence:NativeDocxTextboxGeometryEvidenceV1={items:[{owner:{package_sha256:document.source.package_sha256,part_sha256:'sha256:'+'a'.repeat(64),paragraph_id:p.id,diagnostic_id:'shape:1',anchor:{...anchor},kind:'drawingml',status:'supported',paragraphs:['Rectangle source'],reason:''},geometry:{width_emu:2743200,height_emu:914400,insets_emu:[91440,91440,91440,91440],fill_rgb:'FFF2CC',line_rgb:'204060',line_width_emu:12700,font_family:'DejaVu Sans',font_size_half_points:24,text_rgb:'102030'}}],omitted_count:0}
 return {document,evidence}
}
describe('authored rectangle textbox preview',()=>{
 it('shapes actual supplied font into source bounds and preserves source inputs',()=>{
  const {document,evidence}=fixture(),before=structuredClone({document,evidence}),result=compileNativeDocxTextboxShapeV1(document,evidence,0,font)
  expect(result.status,result.reason).toBe('supported');expect(result.width_millipoints).toBe(216000);expect(result.height_millipoints).toBe(72000);expect(result.line_width_millipoints).toBe(1000);expect(result.paths.length).toBeGreaterThan(4);expect(result.font_sha256).toBe(nativeTextboxFontDigestV1(font))
  expect(decodePaint(document,evidence,0,result,nativeTextboxFontDigestV1(font))).toEqual(result);expect({document,evidence}).toEqual(before)
  expect(compileNativeDocxTextboxShapeV1(document,evidence,0,font)).toEqual(result)
 })
 it('refuses overflow and wrong regular face identity without glyph leakage',()=>{
  for(const mutate of [(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.font_family='Arial'},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.width_emu=200025},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.height_emu=190500}]){
   const {document,evidence}=fixture();mutate(evidence);const result=compileNativeDocxTextboxShapeV1(document,evidence,0,font);expect(result.status).toBe('omitted');expect(result.paths).toEqual([])
  }
  const {document,evidence}=fixture(),bold=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans-Bold.ttf')));expect(compileNativeDocxTextboxShapeV1(document,evidence,0,bold).status).toBe('omitted')
 })
 it('rejects stale/hostile/ambiguous geometry before compiling',()=>{
  for(const mutate of [(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.owner.package_sha256='sha256:'+'f'.repeat(64)},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.width_emu++},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.insets_emu=[2743200,0,0,0]},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.line_rgb='url(secret)'},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.owner.kind='vml'}]){const {document,evidence}=fixture();mutate(evidence);expect(()=>decode(document,evidence)).toThrow()}
  const {document,evidence}=fixture();let invoked=false;Object.defineProperty(evidence.items[0]!.geometry,'font_family',{get(){invoked=true;return 'font'}});expect(()=>decode(document,evidence)).toThrow();expect(invoked).toBe(false)
  let dag:unknown='x';for(let i=0;i<12;i++)dag=[dag,dag,dag];expect(()=>nativeTextboxGeometryPlainData(dag)).toThrow('budget')
 })
 it('rejects helper glyph escapes and wrong uploaded font hash',()=>{
  const {document,evidence}=fixture(),result=compileNativeDocxTextboxShapeV1(document,evidence,0,font),hash=nativeTextboxFontDigestV1(font)
  expect(()=>decodePaint(document,evidence,0,result,'sha256:'+'0'.repeat(64))).toThrow()
  for(const path of ['M -1 0','M 999999999999999999 1','M 1 1 Q 2','M 8000 8000 L NaN 100','M 8000 8000 <script>'])expect(()=>decodePaint(document,evidence,0,{...result,paths:[path]},hash)).toThrow()
 })
})
