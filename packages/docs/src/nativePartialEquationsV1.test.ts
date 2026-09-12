import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {createNativeDocxEquationPreviewsV1 as preview,decodeNativeDocxEquationContextNoticesV1 as decodeNotices,type NativeDocxEquationContextNoticeV1} from './nativePartialEquationsV1.js'
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
 it('validates indexed radical arity and source joins while retaining hidden guards',()=>{
  const {document,resolved,equation}=fixture()
  const tree={kind:'indexed-radical',children:[{kind:'text',text:'x + 1'},{kind:'text',text:'3'}]}
  const indexed={...equation,tree}
  expect(preview(document,resolved,[indexed])[0]!.tree).toEqual(tree)
  for(const children of [[],tree.children.slice(0,1),[...tree.children,tree.children[0]], [{kind:'text',text:'x'},{kind:'text',text:''}], [{kind:'text',text:'x'},{kind:'text',text:'n'.repeat(32769)}], [{kind:'text',text:'x'},{kind:'text',text:'3',href:'javascript:bad'}]])expect(()=>preview(document,resolved,[{...indexed,tree:{...tree,children}}])).toThrow(TypeError)
  expect(()=>preview(document,resolved,[{...indexed,package_sha256:'sha256:'+'0'.repeat(64)}])).toThrow(TypeError)
  expect(()=>preview(document,resolved,[{...indexed,anchor:{...indexed.anchor,xml_sha256:'sha256:'+'0'.repeat(64)}}])).toThrow(TypeError)
  resolved.paragraphs[0]!.paragraph_mark_properties.hidden=true
  expect(preview(document,resolved,[indexed])[0]).toMatchObject({status:'omitted'})
  expect(preview(document,resolved,[indexed])[0]).not.toHaveProperty('tree')
 })
 it('qualifies only exact ignored charset declarations and bounded non-drawing tab notices',()=>{
  const {document,resolved,equation}=fixture(),style=document.passthrough_parts.find(p=>p.part_name==='word/styles.xml')!
  const font={...style,part_name:'word/fontTable.xml',content_type:'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'};document.passthrough_parts.push(font)
  resolved.source_parts.font_table_part=font.part_name;resolved.source_parts.styles_part=style.part_name
  const fa={part_name:font.part_name,path:'/w:fonts[1]/w:font[1]/w:charset[1]',start_byte:10,end_byte:60,xml_sha256:'sha256:'+'a'.repeat(64)},ta={...fa,part_name:style.part_name,path:'/w:styles[1]/w:style[1]/w:pPr[1]/w:tabs[1]'}
  resolved.diagnostics.push({code:'UNMODELED_FONT_METADATA',scope_id:document.document_id,part_name:fa.part_name,path:fa.path,severity:'unsupported',preservation:'preserve-verbatim',message:'Ignored matching only'},{code:'UNMODELED_PARAGRAPH_PROPERTY',scope_id:equation.paragraph_id,part_name:ta.part_name,path:ta.path,severity:'unsupported',preservation:'preserve-verbatim',message:'Tab stops'})
  const base={package_sha256:document.source.package_sha256,part_sha256:font.sha256,diagnostic_origin:'resolved' as const}
  const fn:NativeDocxEquationContextNoticeV1={...base,kind:'ignored-font-matching',anchor:fa,code:'UNMODELED_FONT_METADATA',scope_id:document.document_id,value:'80',character_set:'utf-8'},tn:NativeDocxEquationContextNoticeV1={...base,kind:'unused-paragraph-tab-stops',anchor:ta,code:'UNMODELED_PARAGRAPH_PROPERTY',scope_id:equation.paragraph_id,tab_stops:[{kind:'left',position_twips:709,leader:'none'}]}
  expect(preview(document,resolved,[equation],[fn,tn])[0]!.status).toBe('supported')
  for(const change of [{value:'02'},{character_set:'unknown'},{anchor:{...fa,path:'/w:fonts[1]/w:font[1]/w:charset[1]/extra'}}])expect(()=>decodeNotices(document,resolved,[{...fn,...change}])).toThrow(TypeError)
  for(const stops of [[],[{kind:'bar',position_twips:709,leader:'none'}],[{kind:'left',position_twips:-1,leader:'none'}],[{kind:'left',position_twips:709,leader:'dot'}],[...tn.tab_stops!,...tn.tab_stops!]])expect(()=>decodeNotices(document,resolved,[{...tn,tab_stops:stops}])).toThrow(TypeError)
 })
 it('uses exact same-source notices only for equations while retaining diagnostics and hidden guards',()=>{
  const {document,resolved,equation}=fixture(),part=document.passthrough_parts.find(p=>p.part_name==='word/styles.xml')!
  resolved.source_parts.styles_part=part.part_name
  const anchor={part_name:part.part_name,path:'/w:styles[1]/w:style[1]/w:pPr[1]/w:suppressAutoHyphens[1]',start_byte:100,end_byte:160,xml_sha256:'sha256:'+'a'.repeat(64)}
  resolved.diagnostics.push({code:'UNMODELED_PARAGRAPH_PROPERTY',scope_id:equation.paragraph_id,part_name:part.part_name,path:anchor.path,severity:'unsupported',preservation:'preserve-verbatim',message:'Retained'})
  const notice:NativeDocxEquationContextNoticeV1={kind:'disabled-paragraph-hyphenation',package_sha256:document.source.package_sha256,part_sha256:part.sha256,anchor,diagnostic_origin:'resolved',code:'UNMODELED_PARAGRAPH_PROPERTY',scope_id:equation.paragraph_id,value:'true'}
  const before=structuredClone({document,resolved,notice})
  expect(preview(document,resolved,[equation])[0]!.status).toBe('omitted')
  expect(preview(document,resolved,[equation],[notice])[0]).toMatchObject({status:'supported',context_notice_ids:[expect.any(String)]})
  expect({document,resolved,notice}).toEqual(before)
  for(const change of [{package_sha256:'sha256:'+'0'.repeat(64)},{part_sha256:'sha256:'+'0'.repeat(64)},{anchor:{...anchor,part_name:'word/foreign.xml'}},{anchor:{...anchor,path:anchor.path+'/nested'}},{anchor:{...anchor,end_byte:part.byte_length+1}},{value:'false'},{scope_id:document.document_id},{code:'UNKNOWN'}])expect(()=>decodeNotices(document,resolved,[{...notice,...change}])).toThrow(TypeError)
  expect(()=>decodeNotices(document,resolved,[notice,notice])).toThrow(TypeError)
  resolved.paragraphs[0]!.paragraph_mark_properties.hidden=true
  expect(preview(document,resolved,[equation],[notice])[0]!.status).toBe('omitted')
 })
 it('joins horizontal section notice to its entire original diagnostic anchor',()=>{
  const {document,resolved,equation}=fixture(),section=document.sections[0]!
  const anchor={...section.anchor,path:section.anchor.path+'/w:textDirection[1]'}
  document.unsupported.push({id:'section:direction',code:'UNMODELED_SECTION_PROPERTY',scope_id:section.id,anchor,capability:'sections',preservation:'refuse-mutation',message:'Direction'})
  const notice:NativeDocxEquationContextNoticeV1={kind:'horizontal-section',package_sha256:document.source.package_sha256,part_sha256:'sha256:'+'a'.repeat(64),anchor,diagnostic_origin:'document',diagnostic_id:'section:direction',code:'UNMODELED_SECTION_PROPERTY',scope_id:section.id,value:'lrTb'}
  expect(preview(document,resolved,[equation],[notice])[0]!.status).toBe('supported')
  for(const change of [{value:'tbRl'},{anchor:{...anchor,xml_sha256:'sha256:'+'0'.repeat(64)}},{anchor:{...anchor,start_byte:anchor.start_byte+1}},{diagnostic_id:'wrong'}])expect(()=>decodeNotices(document,resolved,[{...notice,...change}])).toThrow(TypeError)
 })
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
