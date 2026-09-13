import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {createNativeDocxReviewInventoryV1 as project,decodeNativeDocxReviewEvidenceV1 as decode,type NativeDocxReviewEvidenceV1} from './nativeReviewInventoryV1.js'
import {createNativeDocxPartialContentPreviewV1} from './nativePartialContentV1.js'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
const original=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1
const options={policy:'source-review-inventory-v1' as const,read_only:true as const}
function fixture(){
 const document=structuredClone(original),p=document.body.blocks[0]!.paragraph!,r=p.runs[0]!
 p.runs=[r];r.text='Inserted <script>source</script>';r.anchor.path=p.anchor.path+'/w:ins[1]/w:r[1]/w:t[1]'
 const anchor={...p.anchor,path:p.anchor.path+'/w:ins[1]',start_byte:200,end_byte:400}
 document.unsupported=[{id:'review:1',code:'WRAPPED_RUN_MARKUP',capability:'run-structure',scope_id:p.id,anchor,preservation:'refuse-mutation',message:'Revision preserved'}]
 const resolved:NativeDocxResolvedLayoutInputV1={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[{paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}}],runs:[{run_id:r.id,paragraph_id:p.id,applied_paragraph_styles:[],applied_character_styles:[],properties:{}}],tables:[],fonts:[],diagnostics:[]}
 const evidence:NativeDocxReviewEvidenceV1={items:[{package_sha256:document.source.package_sha256,paragraph_id:p.id,diagnostic_id:'review:1',anchor:structuredClone(anchor),kind:'insertion',revision_id:'7',author:'Author <script>',created_at:'2026-09-12',run_ids:[r.id]}],omitted_count:0}
 return {document,resolved,evidence,p,r}
}
describe('explicit native review source inventory',()=>{
 it('shows exact insertion text and authored metadata separately without changing ordinary preview or source diagnostics',()=>{
  const {document,resolved,evidence}=fixture(),before=structuredClone({document,resolved,evidence})
  expect(JSON.stringify(createNativeDocxPartialContentPreviewV1(document,{policy:'source-text-with-omissions-v1',read_only:true},resolved))).not.toContain('Inserted <script>source')
  const out=project(document,options,resolved,evidence)
  expect(out.items[0]).toMatchObject({kind:'insertion',text_status:'qualified-insertion',revision_id:'7',author:'Author <script>',segments:[{kind:'text',text:'Inserted <script>source</script>'}],retained_diagnostic_ids:['review:1']})
  expect(out.source_diagnostics.document).toEqual(document.unsupported);expect({document,resolved,evidence}).toEqual(before)
  expect(()=>project(document,{...options,read_only:false} as unknown as typeof options,resolved,evidence)).toThrow()
  expect(()=>project(document,{...options,extra:true} as typeof options,resolved,evidence)).toThrow()
 })
 it('retains global, ancestor, paragraph and resolved blockers and requires exact inherited visibility',()=>{
  for(const scope of ['global','body','paragraph','run','resolved','hidden','missing']){
   const {document,resolved,evidence,p,r}=fixture()
   if(scope==='hidden')resolved.runs[0]!.properties.hidden=true
   else if(scope==='missing')resolved.runs=[]
   else if(scope==='resolved')resolved.diagnostics=[{code:'UNMODELED',scope_id:p.id,part_name:p.anchor.part_name,path:p.anchor.path,preservation:'preserve-verbatim',severity:'unsupported',message:'Retained'}]
   else{const owner=scope==='body'?document.body:scope==='run'?r:p;document.unsupported.push({id:'blocker',code:'UNKNOWN',capability:'source',scope_id:scope==='global'?document.document_id:owner.id,anchor:owner.anchor,preservation:'refuse-mutation',message:'Retained'})}
   const out=project(document,options,resolved,evidence)
   expect(out.items[0]!.text_status,scope).toBe('omitted');expect(out.items[0]!.segments).toEqual([])
   expect(out.items[0]!.author).toBe('Author <script>');expect(out.source_diagnostics.document).toEqual(document.unsupported)
  }
  const {document,evidence}=fixture();expect(project(document,options,undefined,evidence).items[0]!.text_status).toBe('omitted')
 })
 it('never exposes deletion or moved text and rejects run evidence for those kinds',()=>{
  for(const [kind,tag]of [['deletion','del'],['move-from','moveFrom'],['move-to','moveTo']] as const){
   const {document,resolved,evidence,p,r}=fixture(),item=evidence.items[0]!,diag=document.unsupported[0]!
   item.kind=kind;item.anchor.path=p.anchor.path+'/w:'+tag+'[1]';diag.anchor=structuredClone(item.anchor);diag.code=kind==='move-to'?'WRAPPED_RUN_MARKUP':'UNMODELED_PARAGRAPH_CONTENT';r.anchor.path=item.anchor.path+'/w:r[1]/w:t[1]'
   expect(()=>decode(document,evidence)).toThrow('Only insertion')
   item.run_ids=[];const out=project(document,options,resolved,evidence)
   expect(out.items[0]!.text_status).toBe('omitted');expect(out.items[0]!.segments).toEqual([])
  }
 })
 it('rejects stale, forged, nested, incomplete, reordered and hostile sidecar data',()=>{
  for(const mutate of [
   (e:NativeDocxReviewEvidenceV1)=>e.items[0]!.package_sha256='sha256:'+'f'.repeat(64),
   (e:NativeDocxReviewEvidenceV1)=>e.items[0]!.anchor.xml_sha256='sha256:'+'f'.repeat(64),
   (e:NativeDocxReviewEvidenceV1)=>e.items[0]!.paragraph_id='missing',
   (e:NativeDocxReviewEvidenceV1)=>e.items[0]!.run_ids=['missing'],
   (e:NativeDocxReviewEvidenceV1)=>e.items.push(structuredClone(e.items[0]!)),
   (e:NativeDocxReviewEvidenceV1)=>e.items[0]!.author='a'.repeat(1025),
   (e:NativeDocxReviewEvidenceV1)=>Object.defineProperty(e.items[0]!,'author',{get(){throw Error('getter')}}),
  ]){const {document,evidence}=fixture();mutate(evidence);expect(()=>decode(document,evidence)).toThrow()}
  const {document,evidence,p,r}=fixture();r.anchor.path=p.anchor.path+'/w:ins[1]/w:ins[1]/w:r[1]/w:t[1]';expect(()=>decode(document,evidence)).toThrow('Invalid direct insertion')
  r.anchor.path=p.anchor.path+'/w:ins[1]/w:r[1]/w:t[1]';const second={...r,id:'second',anchor:{...r.anchor,path:p.anchor.path+'/w:ins[1]/w:r[2]/w:t[1]'}};p.runs.push(second)
  expect(()=>decode(document,evidence)).toThrow('incomplete');evidence.items[0]!.run_ids=[second.id,r.id];expect(()=>decode(document,evidence)).toThrow('reordered')
 })
 it('bounds review text and does not reconstruct mixed revision paragraphs',()=>{
  const {document,resolved,evidence,p,r}=fixture();r.text='x'.repeat(32768)
  expect(project(document,options,resolved,evidence).items[0]!.text_status).toBe('omitted')
  r.text='Inserted';document.unsupported.push({id:'deleted',code:'UNMODELED_PARAGRAPH_CONTENT',capability:'run-structure',scope_id:p.id,anchor:{...p.anchor,path:p.anchor.path+'/w:del[1]',start_byte:450,end_byte:600},preservation:'refuse-mutation',message:'Deletion'})
  expect(project(document,options,resolved,evidence).items[0]!.text_status).toBe('omitted')
  expect(decode(document,undefined)).toEqual({items:[],omitted_count:0})
 })
})
