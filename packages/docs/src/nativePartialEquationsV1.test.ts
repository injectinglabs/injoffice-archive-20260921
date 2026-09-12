import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {createNativeDocxEquationPreviewsV1 as preview} from './nativePartialEquationsV1.js'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
function fixture(){
 const document=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1
 const p=document.body.blocks[0]!.paragraph!,anchor={...p.anchor,path:p.anchor.path+'/ns12345678:oMath[1]',start_byte:p.anchor.start_byte+1,end_byte:p.anchor.end_byte-1}
 document.unsupported=[{id:'equation:1',code:'UNMODELED_PARAGRAPH_CONTENT',scope_id:p.id,anchor,capability:'run-structure',preservation:'preserve-verbatim',message:'Opaque equation'}]
 const resolved:NativeDocxResolvedLayoutInputV1={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[{paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}}],runs:[],tables:[],fonts:[],diagnostics:[]}
 const equation={package_sha256:document.source.package_sha256,paragraph_id:p.id,anchor,diagnostic_id:'equation:1',status:'supported',tree:{kind:'fraction',children:[{kind:'text',text:'<script>x</script>'},{kind:'radical',children:[{kind:'text',text:'y'}]}]}}
 return {document,resolved,equation}
}
describe('source-bound read-only equation preview',()=>{
 it('recovers independently qualified equations in separate paragraphs without clearing unrelated blockers',()=>{
  const {document,resolved,equation}=fixture(),p=document.body.blocks[1]!.table!.rows[0]!.cells[0]!.paragraphs[0]!
  const anchor={...p.anchor,path:p.anchor.path+'/ns12345678:oMath[1]',start_byte:p.anchor.start_byte+1,end_byte:p.anchor.end_byte-1}
  const other={...equation,paragraph_id:p.id,diagnostic_id:'equation:2',anchor}
  document.unsupported.push({...document.unsupported[0]!,id:other.diagnostic_id,scope_id:p.id,anchor})
  resolved.paragraphs.push({paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}})
  expect(preview(document,resolved,[equation,other]).map(e=>e.status)).toEqual(['supported','supported'])
  resolved.paragraphs[1]!.paragraph_mark_properties.hidden=true
  expect(preview(document,resolved,[equation,other]).map(e=>e.status)).toEqual(['supported','omitted'])
  document.unsupported.push({...document.unsupported[0]!,id:'unknown',code:'UNKNOWN_VISIBILITY'})
  expect(preview(document,resolved,[equation,other]).map(e=>e.status)).toEqual(['omitted','omitted'])
 })
 it('retains strict source and returns only bounded literal tree data',()=>{
  const {document,resolved,equation}=fixture(),before=structuredClone({document,resolved,equation})
  expect(preview(document,resolved,[equation])[0]).toEqual(equation)
  expect({document,resolved,equation}).toEqual(before)
 })
 it('omits unqualified hidden or unknown surrounding source',()=>{
  const {document,resolved,equation}=fixture();resolved.paragraphs[0]!.paragraph_mark_properties.hidden=true
  expect(preview(document,resolved,[equation])[0]).toMatchObject({status:'omitted'})
  delete resolved.paragraphs[0]!.paragraph_mark_properties.hidden
  document.unsupported.push({...document.unsupported[0]!,id:'unsafe',code:'UNKNOWN_VISIBILITY'})
  expect(preview(document,resolved,[equation])[0]).not.toHaveProperty('tree')
 })
 it('rejects stale or forged source and executable-shaped trees',()=>{
  const {document,resolved,equation}=fixture()
  for(const change of [{package_sha256:'sha256:'+'0'.repeat(64)},{diagnostic_id:'unknown'},{paragraph_id:'unknown'},{anchor:{...equation.anchor,path:equation.anchor.path+'/nested'}},{tree:{kind:'script',text:'bad'}},{tree:{kind:'text',text:'x',href:'javascript:bad'}},{tree:{kind:'fraction',children:[{kind:'text',text:'x'}]}},{tree:{kind:'text',text:'x'.repeat(32769)}}])expect(()=>preview(document,resolved,[{...equation,...change}])).toThrow(TypeError)
  expect(()=>preview(document,resolved,[equation,equation])).toThrow(TypeError)
 })
 it('never invokes tree property getters',()=>{
  const {document,resolved,equation}=fixture();let calls=0
  expect(()=>preview(document,resolved,[{...equation,tree:{get kind(){calls++;return 'text'},text:'x'}}])).toThrow(TypeError)
  expect(calls).toBe(0)
 })
 it('rejects hostile or non-dense arrays before reading any sibling',()=>{
  const {document,resolved,equation}=fixture();let calls=0
  const getter:unknown[]=[];Object.defineProperty(getter,'0',{get(){calls++;return equation},enumerable:true});getter.length=1
  const sparse=new Array(1),custom=[equation];Object.setPrototypeOf(custom,{})
  const symbolic=[equation];Object.defineProperty(symbolic,Symbol('extra'),{value:true})
  const extra=[equation];Object.defineProperty(extra,'other',{value:true})
  for(const array of [getter,sparse,custom,symbolic,extra])expect(()=>preview(document,resolved,array)).toThrow(TypeError)
  const children:unknown[]=[];Object.defineProperty(children,'0',{get(){calls++;return {kind:'text',text:'x'}},enumerable:true});children.length=1
  expect(()=>preview(document,resolved,[{...equation,tree:{kind:'row',children}}])).toThrow(TypeError)
  expect(calls).toBe(0)
 })
})
