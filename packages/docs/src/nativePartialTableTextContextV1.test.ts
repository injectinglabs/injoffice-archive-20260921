import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {decodeNativeDocxPartialTableTextContextsV1 as decode,type NativeDocxPartialTableTextContextV1} from './nativePartialTableTextContextV1.js'
import {createNativeDocxPartialContentPreviewV1 as project} from './nativePartialContentV1.js'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
const original=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1
const options={policy:'source-text-with-omissions-v1' as const,read_only:true as const}
function fixture(){
 const document=structuredClone(original),table=document.body.blocks[1]!.table!,anchor={...table.anchor,path:table.anchor.path+'/w:tblPr[1]/w:tblLook[1]',start_byte:table.anchor.start_byte+1,end_byte:table.anchor.start_byte+2}
 document.unsupported=[{id:'look:1',code:'UNMODELED_TABLE_PROPERTY',scope_id:table.id,anchor,capability:'table-properties',preservation:'refuse-mutation',message:'Look retained'}]
 const resolved:NativeDocxResolvedLayoutInputV1={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part,styles_part:'word/styles.xml'},paragraphs:[],runs:[],tables:[{table_id:table.id,style_id:table.table_style_id}],fonts:[],diagnostics:[{code:'TABLE_STYLE_EFFECTS_PRESERVED',scope_id:table.id,part_name:'word/styles.xml',path:'/w:styles[1]/w:style[1]/w:tblPr[1]',severity:'unsupported',preservation:'preserve-verbatim',message:'Borders retained'}]}
 for(const block of document.body.blocks)for(const p of block.paragraph?[block.paragraph]:block.table!.rows.flatMap(r=>r.cells.flatMap(c=>c.paragraphs))){resolved.paragraphs.push({paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}});for(const r of p.runs)resolved.runs.push({run_id:r.id,paragraph_id:p.id,applied_paragraph_styles:[],applied_character_styles:[],properties:{}})}
 const contexts:NativeDocxPartialTableTextContextV1[]=[{package_sha256:document.source.package_sha256,table_id:table.id,look_diagnostic_id:'look:1',look_anchor:structuredClone(anchor),styles_part:'word/styles.xml',styles_sha256:document.passthrough_parts.find(p=>p.part_name==='word/styles.xml')!.sha256,style_chain:[{style_id:'TableGrid',anchor:{...anchor,part_name:'word/styles.xml',path:'/w:styles[1]/w:style[1]',start_byte:1,end_byte:400}}],resolved_diagnostics:[{code:'TABLE_STYLE_EFFECTS_PRESERVED',scope_id:table.id,part_name:'word/styles.xml',path:'/w:styles[1]/w:style[1]/w:tblPr[1]'}]}]
 return {document,resolved,contexts,table}
}
describe('source-qualified geometry-only table text context',()=>{
 it('recovers owner text with exact same-byte context and retains source diagnostics and geometry omission',()=>{
  const {document,resolved,contexts}=fixture(),before=structuredClone({document,resolved,contexts})
  expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
  const out=project(document,options,resolved,undefined,contexts)
  expect(JSON.stringify(out.blocks)).toContain('Summary');expect(out.omissions.some(o=>o.code==='table')).toBe(true)
  expect(out.retained_nontext_diagnostic_ids).toContain('look:1');expect(out.retained_nontext_diagnostic_ids).toContain('resolved:0:TABLE_STYLE_EFFECTS_PRESERVED')
  expect(out.source_diagnostics).toEqual({document:document.unsupported,resolved:resolved.diagnostics});expect({document,resolved,contexts}).toEqual(before)
 })
 it('retains every other global, table, row, cell, paragraph, run and resolved visibility guard',()=>{
  for(const scope of ['global','table','row','cell','paragraph','run','hidden','missing','resolved']){
   const {document,resolved,contexts,table}=fixture(),row=table.rows[0]!,cell=row.cells[0]!,p=cell.paragraphs[0]!,r=p.runs[0]!
   if(scope==='hidden')resolved.runs.find(x=>x.run_id===r.id)!.properties.hidden=true
   else if(scope==='missing')resolved.runs=resolved.runs.filter(x=>x.run_id!==r.id)
   else if(scope==='resolved')resolved.diagnostics.push({code:'CONDITIONAL_TABLE_STYLE_PRESERVED',scope_id:table.id,part_name:'word/styles.xml',path:'/w:styles[1]/w:style[1]',severity:'unsupported',preservation:'preserve-verbatim',message:'Conditional visibility unqualified'})
   else{const owner=scope==='table'?table:scope==='row'?row:scope==='cell'?cell:scope==='run'?r:p;document.unsupported.push({id:'other',code:'UNMODELED',scope_id:scope==='global'?document.document_id:owner.id,anchor:owner.anchor,capability:'source',preservation:'refuse-mutation',message:'Retained'})}
   const out=project(document,options,resolved,undefined,contexts);expect(JSON.stringify(out.blocks),scope).not.toContain('Summary')
  }
 })
 it('rejects stale source, style-part, look-anchor, chain and resolved diagnostic joins',()=>{
  for(const mutate of [
   (x:NativeDocxPartialTableTextContextV1)=>x.package_sha256='sha256:'+'0'.repeat(64),
   (x:NativeDocxPartialTableTextContextV1)=>x.styles_sha256='sha256:'+'0'.repeat(64),
   (x:NativeDocxPartialTableTextContextV1)=>x.look_anchor.xml_sha256='sha256:'+'0'.repeat(64),
   (x:NativeDocxPartialTableTextContextV1)=>x.style_chain[0]!.style_id='Missing',
   (x:NativeDocxPartialTableTextContextV1)=>x.style_chain[0]!.anchor.end_byte=1000,
   (x:NativeDocxPartialTableTextContextV1)=>x.resolved_diagnostics[0]!.path='/w:styles[1]/w:style[2]/w:tblPr[1]',
   (x:NativeDocxPartialTableTextContextV1)=>x.style_chain.push(structuredClone(x.style_chain[0]!)),
   (x:NativeDocxPartialTableTextContextV1)=>Object.defineProperty(x,'styles_part',{get(){throw new Error('getter')}}),
  ]){const {document,resolved,contexts}=fixture();mutate(contexts[0]!);expect(()=>decode(document,resolved,contexts)).toThrow()}
  const {document,resolved,contexts}=fixture();contexts.push(structuredClone(contexts[0]!));expect(()=>decode(document,resolved,contexts)).toThrow();expect(()=>decode(document,undefined,contexts)).toThrow()
  contexts.pop();resolved.diagnostics=[];expect(()=>decode(document,resolved,contexts)).toThrow('original resolved diagnostic')
 })
 // The producer omits look_diagnostic_id exactly when the extractor proved the
 // conditional selection inactive, so there is no diagnostic left to name. The
 // element bytes are the same evidence and the decoder must admit that form --
 // but only when the document really states no such diagnostic.
 it('admits an inert look that names no diagnostic and rejects one that hides a stated diagnostic',()=>{
  const {document,resolved,contexts}=fixture()
  delete contexts[0]!.look_diagnostic_id
  expect(()=>decode(document,resolved,contexts)).toThrow('omits a look diagnostic')
  document.unsupported=[]
  expect(decode(document,resolved,contexts)[0]!.look_diagnostic_id).toBeUndefined()
  const out=project(document,options,resolved,undefined,contexts)
  expect(JSON.stringify(out.blocks)).toContain('Summary');expect(out.retained_nontext_diagnostic_ids).not.toContain('look:1')
  contexts[0]!.look_anchor.start_byte=0
  expect(()=>decode(document,resolved,contexts)).toThrow('exact source look')
 })
 it('preserves missing styles and continuation omissions and bounds evidence',()=>{
  const {document,resolved,contexts,table}=fixture()
  table.rows[0]!.cells[0]!.vertical_merge='continue';expect(JSON.stringify(project(document,options,resolved,undefined,contexts).blocks)).not.toContain('Summary')
  table.rows[0]!.cells[0]!.vertical_merge='none';resolved.diagnostics.push({code:'MISSING_TABLE_STYLE',scope_id:table.id,severity:'unsupported',preservation:'preserve-verbatim',message:'Missing source style'})
  expect(JSON.stringify(project(document,options,resolved,undefined,contexts).blocks)).not.toContain('Summary')
  expect(()=>decode(document,resolved,Array(65).fill(contexts[0]))).toThrow();expect(decode(document,resolved,undefined)).toEqual([])
 })
})
