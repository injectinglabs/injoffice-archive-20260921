import {decodeNativeDocxDocument,type NativeDocxSourceAnchorV1} from './nativeContract.js'
import {decodeNativeDocxResolvedLayout,type NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
import {decodeNativeDocxNestedTableOmissionsV1,type NativeDocxNestedTableOmissionsV1} from './nativePartialNestedTablesV1.js'
import {decodeNativeDocxPartialTableTextContextsV1} from './nativePartialTableTextContextV1.js'
import {createNativeDocxPartialContentPreviewV1,type NativeDocxPartialParagraphV1} from './nativePartialContentV1.js'
type Owner=NativeDocxNestedTableOmissionsV1['items'][number]
export interface NativeDocxNestedTextEvidenceV1 {owner:Owner;style_id:string;paragraphs:Array<{id:string;anchor:NativeDocxSourceAnchorV1;runs:Array<{id:string;anchor:NativeDocxSourceAnchorV1;text:string}>}>;resolved_layout:NativeDocxResolvedLayoutInputV1}
const fail=():never=>{throw new TypeError('Unqualified nested source text evidence')}
function record(v:unknown,keys:string):Record<string,unknown>{
 if(!v||typeof v!=='object'||Object.getPrototypeOf(v)!==Object.prototype)return fail()
 const expected=keys.split(' '),actual=Reflect.ownKeys(v)
 if(actual.length!==expected.length||actual.some(k=>typeof k!=='string'||!expected.includes(k)))return fail()
 const out:Record<string,unknown>={};for(const k of expected){const d=Object.getOwnPropertyDescriptor(v,k);if(!d||!('value'in d))return fail();out[k]=d.value}return out
}
function array(v:unknown,max:number):unknown[]{if(!Array.isArray(v)||Object.getPrototypeOf(v)!==Array.prototype||v.length>max||Reflect.ownKeys(v).length!==v.length+1)return fail();return Array.from({length:v.length},(_,i)=>{const d=Object.getOwnPropertyDescriptor(v,String(i));return d&&'value'in d?d.value:fail()})}
function anchor(v:unknown,parent:NativeDocxSourceAnchorV1,suffix:RegExp):NativeDocxSourceAnchorV1{
 const a=record(v,'part_name path start_byte end_byte xml_sha256')
 if(a.part_name!==parent.part_name||typeof a.path!=='string'||!a.path.startsWith(parent.path)||!suffix.test(a.path.slice(parent.path.length))||!Number.isSafeInteger(a.start_byte)||!Number.isSafeInteger(a.end_byte)||Number(a.start_byte)<parent.start_byte||Number(a.end_byte)>parent.end_byte||Number(a.end_byte)<=Number(a.start_byte)||typeof a.xml_sha256!=='string'||!/^sha256:[a-f0-9]{64}$/.test(a.xml_sha256))return fail()
 return a as unknown as NativeDocxSourceAnchorV1
}
/** Same-byte producer evidence only; auxiliary XML is not reparsed by this decoder. */
export function decodeNativeDocxNestedTextEvidenceV1(source:unknown,resolved:unknown,nestedInput:unknown,contextInput:unknown,input:unknown):NativeDocxNestedTextEvidenceV1[]{
 const d=decodeNativeDocxDocument(source),l=decodeNativeDocxResolvedLayout(resolved)
 if(!d.ok||!l.ok||l.value.document_id!==d.value.document_id||l.value.revision!==d.value.revision||l.value.source_parts.main_part!==d.value.source.main_part)return fail()
 const nested=decodeNativeDocxNestedTableOmissionsV1(d.value,nestedInput),contexts=decodeNativeDocxPartialTableTextContextsV1(d.value,l.value,contextInput)
 if(input===undefined)return []
 const seen=new Set<string>();let units=0
 return array(input,32).map(raw=>{
  const v=record(raw,'owner style_id paragraphs resolved_layout'),o=record(v.owner,'package_sha256 part_sha256 table_id cell_id diagnostic_id anchor'),a=record(o.anchor,'part_name path start_byte end_byte xml_sha256')
  const owner=nested.items.find(n=>n.diagnostic_id===o.diagnostic_id)
  if(!owner||seen.has(owner.diagnostic_id)||Object.entries(owner).some(([k,x])=>k==='anchor'?Object.entries(x).some(([k,x])=>a[k]!==x):o[k]!==x))return fail()
  const context=contexts.find(c=>c.table_id===owner.table_id)
  if(!context||v.style_id!==context.style_chain.at(-1)?.style_id)return fail()
  const layout=decodeNativeDocxResolvedLayout(v.resolved_layout)
  if(!layout.ok||layout.value.document_id!==l.value.document_id||layout.value.revision!==l.value.revision||JSON.stringify(layout.value.source_parts)!==JSON.stringify(l.value.source_parts)||layout.value.tables.length||layout.value.fonts.length||layout.value.diagnostics.length||layout.value.numbering_source)return fail()
  const ids=new Set<string>(),runIDs:string[]=[],paragraphIDs:string[]=[];let previousEnd=owner.anchor.start_byte
  const ps=array(v.paragraphs,16).map(raw=>{
   const p=record(raw,'id anchor runs'),pa=anchor(p.anchor,owner.anchor,/^\/w:tr\[1\]\/w:tc\[1\]\/w:p\[[1-9][0-9]*\]$/)
   if(typeof p.id!=='string'||!/^nested-paragraph:[a-f0-9]{24}$/.test(p.id)||ids.has(p.id)||pa.start_byte<previousEnd)return fail()
   previousEnd=pa.end_byte;ids.add(p.id);paragraphIDs.push(p.id)
   const rp=layout.value.paragraphs.find(r=>r.paragraph_id===p.id)
   if(!rp||rp.numbering||rp.paragraph_mark_properties.hidden)return fail()
   let runEnd=pa.start_byte
   const runs=array(p.runs,32).map(raw=>{
    const r=record(raw,'id anchor text'),ra=anchor(r.anchor,pa,/^\/w:r\[[1-9][0-9]*\]\/w:t\[1\]$/)
    if(typeof r.id!=='string'||!/^nested-run:[a-f0-9]{24}$/.test(r.id)||ids.has(r.id)||ra.start_byte<runEnd||typeof r.text!=='string'||/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(r.text)||(units+=r.text.length)>32768)return fail()
    runEnd=ra.end_byte;ids.add(r.id);runIDs.push(r.id)
    const rr=layout.value.runs.find(x=>x.run_id===r.id)
    if(!rr||rr.paragraph_id!==p.id||rr.properties.hidden)return fail()
    return {id:r.id,anchor:ra,text:r.text}
   })
   return {id:p.id,anchor:pa,runs}
  })
  if(!ps.length||JSON.stringify(paragraphIDs)!==JSON.stringify(layout.value.paragraphs.map(p=>p.paragraph_id))||JSON.stringify(runIDs)!==JSON.stringify(layout.value.runs.map(r=>r.run_id)))return fail()
  seen.add(owner.diagnostic_id);return {owner:structuredClone(owner),style_id:v.style_id as string,paragraphs:ps,resolved_layout:structuredClone(layout.value)}
 })
}
export interface NativeDocxNestedTextInventoryV1 {policy:'source-nested-table-text-v1';read_only:true;source:{document_id:string;revision:string;package_sha256:string};items:Array<{owner:Owner;paragraphs:NativeDocxPartialParagraphV1[]}>;omitted_count:number;source_diagnostics:ReturnType<typeof createNativeDocxPartialContentPreviewV1>['source_diagnostics']}
export function createNativeDocxNestedTextInventoryV1(source:unknown,options:{policy:'source-nested-table-text-v1';read_only:true},resolved:unknown,nested:unknown,contexts:unknown,evidence:unknown):NativeDocxNestedTextInventoryV1{
 const opts=record(options,'policy read_only');if(opts.policy!=='source-nested-table-text-v1'||opts.read_only!==true)return fail()
 const qualified=decodeNativeDocxNestedTextEvidenceV1(source,resolved,nested,contexts,evidence)
 const outer=createNativeDocxPartialContentPreviewV1(source,{policy:'source-text-with-omissions-v1',read_only:true},resolved,nested,contexts)
 // A nested omission is reachable only after every original global, table,
 // row and owner-cell diagnostic/continuation gate has passed.
 const reachable=new Set(outer.omissions.filter(o=>o.code==='nested-table').map(o=>o.source.scope_id))
 const items=qualified.filter(q=>reachable.has(q.owner.diagnostic_id)).map(q=>({owner:q.owner,paragraphs:q.paragraphs.map(p=>({kind:'paragraph' as const,source:{scope_id:p.id,anchor:p.anchor},segments:p.runs.map(r=>({kind:'text' as const,source:{scope_id:r.id,anchor:r.anchor},text:r.text}))}))}))
 const inventory=decodeNativeDocxNestedTableOmissionsV1(source,nested)
 return {policy:'source-nested-table-text-v1',read_only:true,source:outer.source,items,omitted_count:inventory.items.length+inventory.omitted_count-items.length,source_diagnostics:outer.source_diagnostics}
}
