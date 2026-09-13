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

function multiline(count=2){
 const f=fixture(),item=f.evidence.items[0]!,g=item.geometry!,a=item.owner.anchor
 a.end_byte=390;f.document.unsupported[0].anchor.end_byte=390
 const anchor=(path:string,start:number,end:number)=>({...a,path,start_byte:start,end_byte:end})
 const p=anchor(a.path+'/wps:wsp[1]/wps:txbx[1]/w:txbxContent[1]/w:p[1]',210,380),r=anchor(p.path+'/w:r[1]',240,370)
 const texts=Array.from({length:count},(_,i)=>['First authored line','Second authored line','Final authored line'][i]!)
 item.owner.paragraphs=[texts.join('\n')];g.font_size_half_points=20
 let offset=0
 item.hard_break_layout={paragraph_anchor:p,run_anchor:r,spacing_anchor:anchor(p.path+'/w:pPr[1]/w:spacing[1]',215,230),line_step_twips:360,lines:texts.map((text,i)=>{const start=offset;offset+=text.length+1;return {ordinal:i,text_anchor:anchor(r.path+`/w:t[${i+1}]`,250+i*30,260+i*30),break_before_anchor:i?anchor(r.path+`/w:br[${i}]`,240+i*30,245+i*30):null,start_utf16:start,end_utf16:offset-1}})}
 return f
}
describe('authored hard-break rectangles',()=>{
 it('shapes two and three real-font lines with exact centered baselines and source coverage',()=>{
  for(const count of [2,3]){
   const {document,evidence}=multiline(count),before=structuredClone(evidence),result=compileNativeDocxTextboxShapeV1(document,evidence,0,font)
   expect(result.status,result.reason).toBe('supported');expect(result.line_layout?.natural_height_millipoints).toBe(11640);expect(result.line_layout?.line_gap_millipoints).toBe(0)
   expect(result.line_layout?.lines.map(l=>l.baseline_millipoints)).toEqual([19662,37662,55662].slice(0,count));expect(result.line_layout?.lines.at(-1)?.path_end).toBe(result.paths.length)
   expect(compileNativeDocxTextboxShapeV1(document,evidence,0,font)).toEqual(result);expect(evidence).toEqual(before)
  }
 })
 it('refuses odd centered metrics, full-stack overflow and late-line overflow without partial paint',()=>{
  for(const mutate of [(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.font_size_half_points=24},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.geometry!.height_emu=868553},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.hard_break_layout!.line_step_twips=100}]){
   const {document,evidence}=multiline(3);mutate(evidence);const result=compileNativeDocxTextboxShapeV1(document,evidence,0,font);expect(result.status).toBe('omitted');expect(result.paths).toEqual([]);expect(result.line_layout).toBeUndefined()
  }
  const {document,evidence}=multiline(3);evidence.items[0]!.geometry!.height_emu=868680;expect(compileNativeDocxTextboxShapeV1(document,evidence,0,font).status).toBe('supported')
 })
 it('refuses fabricated, reordered or incomplete source line evidence and legacy newlines',()=>{
  for(const mutate of [(e:NativeDocxTextboxGeometryEvidenceV1)=>{delete e.items[0]!.hard_break_layout},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.hard_break_layout!.lines[1]!.ordinal=0},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.hard_break_layout!.lines[1]!.start_utf16--},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.hard_break_layout!.lines[1]!.break_before_anchor=null},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.hard_break_layout!.run_anchor.path+='/w:r[2]'},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.hard_break_layout!.lines[1]!.text_anchor.xml_sha256='bad'},(e:NativeDocxTextboxGeometryEvidenceV1)=>{e.items[0]!.hard_break_layout!.lines.pop()}]){const {document,evidence}=multiline();mutate(evidence);expect(()=>decode(document,evidence)).toThrow()}
 })
 it('retains preserved spaces and refuses late-line glyph overhang without earlier paths',()=>{
  const {document,evidence}=multiline(3),item=evidence.items[0]!,lines=item.hard_break_layout!.lines
  item.owner.paragraphs=[item.owner.paragraphs[0]!.replace('Final authored line','Third authored line')]
  const refused=compileNativeDocxTextboxShapeV1(document,evidence,0,font);expect(refused.reason).toBe('glyph-overflow');expect(refused.paths).toEqual([]);expect(refused.line_layout).toBeUndefined()
  item.owner.paragraphs=[item.owner.paragraphs[0]!.replace('Third authored line',' '.repeat('Third authored line'.length))]
  const result=compileNativeDocxTextboxShapeV1(document,evidence,0,font);expect(result.status,result.reason).toBe('supported');const last=result.line_layout!.lines[2]!;expect(last.path_start).toBe(last.path_end);expect(lines[2]!.end_utf16).toBe(item.owner.paragraphs[0]!.length)
 })
 it('rejects missing/duplicated/reordered line paint and paths moved to another line box',()=>{
  const {document,evidence}=multiline(),paint=compileNativeDocxTextboxShapeV1(document,evidence,0,font),hash=nativeTextboxFontDigestV1(font)
  const corruptions=[(p:typeof paint)=>{p.line_layout!.lines.reverse()},(p:typeof paint)=>{p.line_layout!.lines[1]!.baseline_millipoints++},(p:typeof paint)=>{p.paths=[];for(const l of p.line_layout!.lines)l.path_start=l.path_end=0},(p:typeof paint)=>{p.paths[0]=''},(p:typeof paint)=>{p.paths[0]=' , '},(p:typeof paint)=>{p.paths[0]='M 8000 8000 Z'},(p:typeof paint)=>{p.paths[0]='M 8000 30000 L 8001 30001'},(p:typeof paint)=>{p.line_layout!.lines[0]!.path_end=0},(p:typeof paint)=>{delete p.line_layout}]
  for(const mutate of corruptions){const p=structuredClone(paint);mutate(p);expect(()=>decodePaint(document,evidence,0,p,hash)).toThrow()}
 })
})

function wrapping(text='First authored line Second authored line Final authored line',width=1270000){
 const f=fixture(),item=f.evidence.items[0]!,a=item.owner.anchor,g=item.geometry!
 const anchor=(path:string,start:number,end:number)=>({...a,path,start_byte:start,end_byte:end})
 const shape=a.path+'/wps:wsp[1]',p=anchor(shape+'/wps:txbx[1]/w:txbxContent[1]/w:p[1]',210,340),r=anchor(p.path+'/w:r[1]',240,330)
 item.owner.paragraphs=[text];g.font_size_half_points=20;g.width_emu=width;g.height_emu=1828800;g.text_wrap='square'
 item.wrap_layout={policy:'ascii-space-greedy-v1',body_properties_anchor:anchor(shape+'/wps:bodyPr[1]',350,380),paragraph_anchor:p,run_anchor:r,text_anchor:anchor(r.path+'/w:t[1]',250,320),spacing_anchor:anchor(p.path+'/w:pPr[1]/w:spacing[1]',215,230),line_step_twips:360}
 return f
}
describe('automatic ASCII rectangle wrapping',()=>{
 it('uses complete real-font cluster coverage and changes lines with width',()=>{
  const counts=[]
  for(const width of [1270000,1905000]){
   const {document,evidence}=wrapping(undefined,width),before=structuredClone(evidence),p=compileNativeDocxTextboxShapeV1(document,evidence,0,font)
   expect(p.status,p.reason).toBe('supported');expect(p.wrap_paint).toBeDefined();expect(p.line_layout).toBeUndefined();counts.push(p.wrap_paint!.lines.length)
   expect(p.wrap_paint!.lines.map(l=>evidence.items[0]!.owner.paragraphs[0]!.slice(l.start_utf16,l.end_utf16)).join('')).toBe(evidence.items[0]!.owner.paragraphs[0]);expect(p.wrap_paint!.lines.at(-1)!.path_end).toBe(p.paths.length)
   expect(compileNativeDocxTextboxShapeV1(document,evidence,0,font)).toEqual(p);expect(evidence).toEqual(before)
  }
  expect(counts[0]).toBeGreaterThan(counts[1]!)
 })
 it('retains real font ligature clusters inside wrapped words',()=>{
  const {document,evidence}=wrapping('office office office office'),p=compileNativeDocxTextboxShapeV1(document,evidence,0,font)
  expect(p.status,p.reason).toBe('supported');expect(p.wrap_paint!.clusters.some(c=>c.end_utf16-c.start_utf16>1)).toBe(true)
  for(const line of p.wrap_paint!.lines)expect(evidence.items[0]!.owner.paragraphs[0]!.slice(line.start_utf16,line.end_utf16)).toMatch(/^office(?: office)* ?$/)
 })
 it('refuses unbreakable words, odd centering, stack overflow and invalid source policies',()=>{
  for(const f of [wrapping('M'.repeat(100)),wrapping()]){if(f.evidence.items[0]!.owner.paragraphs[0]!.startsWith('First'))f.evidence.items[0]!.geometry!.font_size_half_points=24;const p=compileNativeDocxTextboxShapeV1(f.document,f.evidence,0,font);expect(p.status).toBe('omitted');expect(p.paths).toEqual([]);expect(p.wrap_paint).toBeUndefined()}
  const f=wrapping();f.evidence.items[0]!.geometry!.height_emu=400050;expect(compileNativeDocxTextboxShapeV1(f.document,f.evidence,0,font).status).toBe('omitted')
  for(const text of [' edge','edge ','two  spaces','punctuation.','soft\nhard','éclair']){const f=wrapping(text);expect(()=>decode(f.document,f.evidence)).toThrow()}
  for(const field of ['geometry','evidence']){const f=wrapping();if(field==='geometry')delete f.evidence.items[0]!.geometry!.text_wrap;else delete f.evidence.items[0]!.wrap_layout;expect(()=>decode(f.document,f.evidence)).toThrow()}
  const h=multiline();h.evidence.items[0]!.wrap_layout=wrapping().evidence.items[0]!.wrap_layout;expect(()=>decode(h.document,h.evidence)).toThrow()
 })
 it('rejects missing source units, fabricated clusters and non-greedy or missing line paint',()=>{
  const {document,evidence}=wrapping(),result=compileNativeDocxTextboxShapeV1(document,evidence,0,font),hash=nativeTextboxFontDigestV1(font);expect(result.status,result.reason).toBe('supported')
  const mutations=[(p:typeof result)=>{p.wrap_paint!.clusters[0]!.start_utf16++},(p:typeof result)=>{p.wrap_paint!.clusters.pop()},(p:typeof result)=>{p.wrap_paint!.clusters[0]!.glyph_start++},(p:typeof result)=>{p.wrap_paint!.clusters[0]!.advance_millipoints=-1},(p:typeof result)=>{p.wrap_paint!.lines[0]!.end_utf16--},(p:typeof result)=>{p.wrap_paint!.lines.reverse()},(p:typeof result)=>{p.wrap_paint!.lines[1]!.baseline_millipoints++},(p:typeof result)=>{p.wrap_paint!.lines[0]!.path_end=0},(p:typeof result)=>{const n=p.wrap_paint!.clusters[p.wrap_paint!.lines[0]!.cluster_end-1]!.glyph_end+1;p.paths=Array(n).fill('M 8000 8000 L 8001 8001');p.wrap_paint!.lines[0]!.path_end=n},(p:typeof result)=>{delete p.wrap_paint}]
  for(const mutate of mutations){const p=structuredClone(result);mutate(p);expect(()=>decodePaint(document,evidence,0,p,hash)).toThrow()}
 })
})
