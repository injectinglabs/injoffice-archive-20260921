import {decodeNativeDocxDocument,type NativeDocxSourceAnchorV1,type NativeDocxDocumentV1} from './nativeContract.js'
import {decodeNativeDocxResolvedLayout,type NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'

export const DOCX_PARTIAL_CONTENT_POLICY='source-text-with-omissions-v1' as const
export const DOCX_PARTIAL_CONTENT_LIMITS={bodyBlocks:200,textUnits:100_000} as const
export type NativeDocxPartialOmissionCode='unsupported-source'|'unqualified-text-visibility'|'hidden-text'|'drawing'|'field'|'reference'|'control'|'table'|'nonbody-story'|'block-limit'|'text-limit'
export interface NativeDocxPartialSourceV1 {scope_id:string;anchor:NativeDocxSourceAnchorV1}
export type NativeDocxPartialSegmentV1=
 | {kind:'text';source:NativeDocxPartialSourceV1;text:string}
 | {kind:'omission';source:NativeDocxPartialSourceV1;code:NativeDocxPartialOmissionCode;count:number;diagnostic_ids:string[]}
export interface NativeDocxPartialContentV1 {
 protocol:'injoffice.docx.partial-content';version:1;policy:typeof DOCX_PARTIAL_CONTENT_POLICY;read_only:true;fidelity:'partial-source-content';pagination:'not-produced'
 source:{document_id:string;revision:string;package_sha256:string}
 blocks:Array<{kind:'paragraph';source:NativeDocxPartialSourceV1;segments:NativeDocxPartialSegmentV1[]}|Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>>
 omissions:Array<Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>>
 source_diagnostics:{document:NativeDocxDocumentV1['unsupported'];resolved:NativeDocxResolvedLayoutInputV1['diagnostics']}
 coverage:{body_blocks:number;visited_body_blocks:number;projected_text_runs:number;omitted_units:number;layout_present:boolean}
 warnings:string[]
}

/** Explicit read-only source-content projection, never input to native paint or
 * mutations. Without joined layout, only the omission inventory is available:
 * inherited visibility cannot safely be inferred from direct run properties.
 * Source hashes are native extractor evidence, not re-hashed package bytes. */
export function createNativeDocxPartialContentPreviewV1(value:unknown,options:{policy:typeof DOCX_PARTIAL_CONTENT_POLICY;read_only:true},resolvedValue?:unknown):NativeDocxPartialContentV1 {
 let validOptions=false
 try{
  if(options&&Object.getPrototypeOf(options)===Object.prototype&&Reflect.ownKeys(options).length===2){
   const descriptors=Object.getOwnPropertyDescriptors(options)
   validOptions=!!descriptors.policy&&'value'in descriptors.policy&&descriptors.policy.value===DOCX_PARTIAL_CONTENT_POLICY&&!!descriptors.read_only&&'value'in descriptors.read_only&&descriptors.read_only.value===true
  }
 }catch{/* Hostile options cannot authorize projection. */}
 if(!validOptions)throw new TypeError('Explicit read-only partial-content policy required')
 const decoded=decodeNativeDocxDocument(value)
 if(!decoded.ok)throw new TypeError('Invalid native document; partial preview cannot bypass integrity validation')
 const document=structuredClone(decoded.value)
 let resolved:NativeDocxResolvedLayoutInputV1|undefined
 if(resolvedValue!==undefined){
  const layout=decodeNativeDocxResolvedLayout(resolvedValue)
  if(!layout.ok||layout.value.document_id!==document.document_id||layout.value.revision!==document.revision||layout.value.source_parts.main_part!==document.source.main_part)throw new TypeError('Partial content requires an exact native layout identity join')
  resolved=structuredClone(layout.value)
 }
 const sources=new Map<string,NativeDocxSourceAnchorV1>()
 const put=(id:string,anchor:NativeDocxSourceAnchorV1)=>sources.set(id,anchor)
 for(const story of [document.body,...document.headers,...document.footers,...document.notes,...document.comment_stories]){
  put(story.id,story.anchor)
  for(const block of story.blocks){
   if(block.paragraph){put(block.paragraph.id,block.paragraph.anchor);for(const run of block.paragraph.runs)put(run.id,run.anchor)}
   if(block.table){put(block.table.id,block.table.anchor);for(const row of block.table.rows){put(row.id,row.anchor);for(const cell of row.cells){put(cell.id,cell.anchor);for(const paragraph of cell.paragraphs){put(paragraph.id,paragraph.anchor);for(const run of paragraph.runs)put(run.id,run.anchor)}}}}
  }
 }
 const blockers=new Map<string,string[]>(),global:string[]=[]
 const add=(scope:string,id:string)=>{if(!sources.has(scope)){global.push(id);return}const list=blockers.get(scope)??[];list.push(id);blockers.set(scope,list)}
 for(const diagnostic of document.unsupported){
  const owner=sources.get(diagnostic.scope_id)
  if(owner&&diagnostic.anchor&&(owner.part_name!==diagnostic.anchor.part_name||!diagnostic.anchor.path.startsWith(owner.path+'/')&&diagnostic.anchor.path!==owner.path)){global.push(diagnostic.id);continue}
  add(diagnostic.scope_id,diagnostic.id)
 }
 for(const [index,diagnostic]of (resolved?.diagnostics??[]).entries()){
  const id=`resolved:${index}:${diagnostic.code}`,owner=sources.get(diagnostic.scope_id)
  if(owner&&(diagnostic.part_name!==undefined&&diagnostic.part_name!==owner.part_name||diagnostic.path!==undefined&&diagnostic.path!==owner.path&&!diagnostic.path.startsWith(owner.path+'/'))){global.push(id);continue}
  add(diagnostic.scope_id,id)
 }
 const paragraphs=new Map(resolved?.paragraphs.map(p=>[p.paragraph_id,p])),runs=new Map(resolved?.runs.map(r=>[r.run_id,r]))
 const omissions:NativeDocxPartialContentV1['omissions']=[],blocks:NativeDocxPartialContentV1['blocks']=[]
 let textUnits=0,textRuns=0
 const source=(scope_id:string,anchor:NativeDocxSourceAnchorV1)=>({scope_id,anchor:structuredClone(anchor)})
 const omit=(s:NativeDocxPartialSourceV1,code:NativeDocxPartialOmissionCode,count=1,diagnostic_ids:string[]=[]):Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>=>{
  const result={kind:'omission' as const,source:s,code,count,diagnostic_ids:[...diagnostic_ids]};omissions.push(result);return result
 }
 const inherited=[...global,...(blockers.get(document.body.id)??[])]
 // One body-scoped placeholder avoids multiplying a large global diagnostic
 // inventory by every block; original diagnostics are retained exactly once.
 if(inherited.length)blocks.push(omit(source(document.body.id,document.body.anchor),'unsupported-source',document.body.blocks.length,inherited))
 for(const block of inherited.length?[]:document.body.blocks.slice(0,DOCX_PARTIAL_CONTENT_LIMITS.bodyBlocks)){
  const node=block.paragraph??block.table
  if(!node)throw new TypeError('Partial preview cannot invent unsupported block ownership')
  const s=source(node.id,node.anchor),diagnostics=[...inherited,...(blockers.get(node.id)??[])]
  if(diagnostics.length){blocks.push(omit(s,'unsupported-source',1,diagnostics));continue}
  if(block.table){blocks.push(omit(s,'table'));continue}
  const paragraph=block.paragraph!,segments:NativeDocxPartialSegmentV1[]=[],p=paragraphs.get(paragraph.id)
  for(const run of paragraph.runs){
   const rs=source(run.id,run.anchor),r=runs.get(run.id),diagnostics=blockers.get(run.id)??[]
   if(diagnostics.length){segments.push(omit(rs,'unsupported-source',1,diagnostics));continue}
   if(run.properties?.hidden||r?.properties.hidden){segments.push(omit(rs,'hidden-text'));continue}
   if(run.page_field){segments.push(omit(rs,'field'));continue}
   if(run.kind!=='text'){segments.push(omit(rs,run.kind==='drawing'?'drawing':run.kind==='reference'?'reference':'control'));continue}
   if(!p||!r||r.paragraph_id!==paragraph.id){segments.push(omit(rs,'unqualified-text-visibility'));continue}
   const text=run.text??''
   if(textUnits+text.length>DOCX_PARTIAL_CONTENT_LIMITS.textUnits){segments.push(omit(rs,'text-limit'));continue}
   textUnits+=text.length;textRuns++;segments.push({kind:'text',source:rs,text})
  }
  blocks.push({kind:'paragraph',source:s,segments})
 }
 const remaining=document.body.blocks.length-DOCX_PARTIAL_CONTENT_LIMITS.bodyBlocks
 if(remaining>0&&!inherited.length)blocks.push(omit(source(document.body.id,document.body.anchor),'block-limit',remaining))
 for(const story of [...document.headers,...document.footers,...document.notes,...document.comment_stories])blocks.push(omit(source(story.id,story.anchor),'nonbody-story',1,blockers.get(story.id)??[]))
 return {protocol:'injoffice.docx.partial-content',version:1,policy:DOCX_PARTIAL_CONTENT_POLICY,read_only:true,fidelity:'partial-source-content',pagination:'not-produced',source:{document_id:document.document_id,revision:document.revision,package_sha256:document.source.package_sha256},blocks,omissions,source_diagnostics:{document:document.unsupported,resolved:resolved?.diagnostics??[]},coverage:{body_blocks:document.body.blocks.length,visited_body_blocks:inherited.length?0:Math.min(document.body.blocks.length,DOCX_PARTIAL_CONTENT_LIMITS.bodyBlocks),projected_text_runs:textRuns,omitted_units:omissions.reduce((n,o)=>n+o.count,0),layout_present:resolved!==undefined},warnings:['Read-only extracted source content, not document pagination. Formatting, list markers, layout and nonbody stories are not reconstructed. Every omitted unit has a source-bound placeholder.',...(resolved?[]:['No resolved layout was supplied: text visibility is unqualified, so this result contains an omission inventory rather than text.'])]}
}
