import {decodeNativeDocxDocument,type NativeDocxSourceAnchorV1,type NativeDocxDocumentV1,type NativeDocxParagraphV1,type NativeDocxRunV1} from './nativeContract.js'
import {decodeNativeDocxResolvedLayout,type NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
import {isRenderNeutralLayoutDiagnostic} from './nativeRenderDiagnostics.js'
import {decodeNativeDocxNestedTableOmissionsV1,type NativeDocxNestedTableOmissionsV1} from './nativePartialNestedTablesV1.js'

export const DOCX_PARTIAL_CONTENT_POLICY='source-text-with-omissions-v1' as const
export const DOCX_PARTIAL_CONTENT_LIMITS={bodyBlocks:200,tableCells:200,textUnits:100_000} as const
export type NativeDocxPartialOmissionCode='unsupported-source'|'unqualified-text-visibility'|'hidden-text'|'drawing'|'field'|'reference'|'control'|'table'|'nonbody-story'|'block-limit'|'text-limit'|'merged-cell'|'cell-limit'|'nested-table'|'nested-table-limit'
export interface NativeDocxPartialSourceV1 {scope_id:string;anchor:NativeDocxSourceAnchorV1}
export type NativeDocxPartialSegmentV1=
 | {kind:'text';source:NativeDocxPartialSourceV1;text:string}
 | {kind:'alternative-text';source:NativeDocxPartialSourceV1;text:string;label:'authored-drawing-description'}
 | {kind:'omission';source:NativeDocxPartialSourceV1;code:NativeDocxPartialOmissionCode;count:number;diagnostic_ids:string[]}
export type NativeDocxPartialParagraphV1={kind:'paragraph';source:NativeDocxPartialSourceV1;segments:NativeDocxPartialSegmentV1[]}
export type NativeDocxPartialCellV1={kind:'cell-source';source:NativeDocxPartialSourceV1;row_ordinal:number;cell_ordinal:number;source_merge?:{grid_span:number;vertical_merge:'none'|'restart'|'continue'};paragraphs:Array<NativeDocxPartialParagraphV1|Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>>}
export interface NativeDocxPartialContentV1 {
 protocol:'injoffice.docx.partial-content';version:1;policy:typeof DOCX_PARTIAL_CONTENT_POLICY;read_only:true;fidelity:'partial-source-content';pagination:'not-produced'
 source:{document_id:string;revision:string;package_sha256:string}
 blocks:Array<NativeDocxPartialParagraphV1|{kind:'table-source';source:NativeDocxPartialSourceV1;source_cell_count:number;cells:NativeDocxPartialCellV1[]}|Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>>
 omissions:Array<Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>>
 source_diagnostics:{document:NativeDocxDocumentV1['unsupported'];resolved:NativeDocxResolvedLayoutInputV1['diagnostics']}
 retained_nontext_diagnostic_ids:string[]
 nested_table_omissions?:NativeDocxNestedTableOmissionsV1
 coverage:{body_blocks:number;visited_body_blocks:number;projected_text_runs:number;omitted_units:number;layout_present:boolean}
 warnings:string[]
}

/** Explicit read-only source-content projection, never input to native paint or
 * mutations. Without joined layout, only the omission inventory is available:
 * inherited visibility cannot safely be inferred from direct run properties.
 * Source hashes are native extractor evidence, not re-hashed package bytes. */
export function createNativeDocxPartialContentPreviewV1(value:unknown,options:{policy:typeof DOCX_PARTIAL_CONTENT_POLICY;read_only:true},resolvedValue?:unknown,nestedEvidence?:unknown):NativeDocxPartialContentV1 {
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
 const nested=decodeNativeDocxNestedTableOmissionsV1(document,nestedEvidence),nestedIDs=new Set(nested.items.map(n=>n.diagnostic_id))
 let resolved:NativeDocxResolvedLayoutInputV1|undefined
 if(resolvedValue!==undefined){
  const layout=decodeNativeDocxResolvedLayout(resolvedValue)
  if(!layout.ok||layout.value.document_id!==document.document_id||layout.value.revision!==document.revision||layout.value.source_parts.main_part!==document.source.main_part)throw new TypeError('Partial content requires an exact native layout identity join')
  resolved=structuredClone(layout.value)
 }
 const sources=new Map<string,NativeDocxSourceAnchorV1>()
 const put=(id:string,anchor:NativeDocxSourceAnchorV1)=>sources.set(id,anchor)
 const putRun=(run:NativeDocxRunV1)=>{put(run.id,run.anchor);if(run.drawing)put(run.drawing.id,run.drawing.anchor)}
 for(const story of [document.body,...document.headers,...document.footers,...document.notes,...document.comment_stories]){
  put(story.id,story.anchor)
  for(const block of story.blocks){
   if(block.paragraph){put(block.paragraph.id,block.paragraph.anchor);for(const run of block.paragraph.runs)putRun(run)}
   if(block.table){put(block.table.id,block.table.anchor);for(const row of block.table.rows){put(row.id,row.anchor);for(const cell of row.cells){put(cell.id,cell.anchor);for(const paragraph of cell.paragraphs){put(paragraph.id,paragraph.anchor);for(const run of paragraph.runs)putRun(run)}}}}
  }
 }
 const blockers=new Map<string,string[]>(),global:string[]=[],retainedNontext:string[]=[]
 const nontext=(diagnostic:NativeDocxResolvedLayoutInputV1['diagnostics'][number]):boolean=>{
  if(!resolved)return false
  // Plain Unicode recovery does not choose fonts or qualify painted glyphs.
  // Retain these precise extractor facts without granting rendering authority.
  if(['LATENT_STYLE_BEHAVIOR_PRESERVED','EMPTY_NUMBERING_STYLE_PRESERVED','FONT_MATCHING_METADATA_PRESERVED'].includes(diagnostic.code)&&isRenderNeutralLayoutDiagnostic(diagnostic,resolved))return true
  if(diagnostic.code!=='TABLE_STYLE_EFFECTS_PRESERVED'||diagnostic.severity!=='unsupported'||diagnostic.preservation!=='preserve-verbatim')return false
  const entry=resolved.tables.find(t=>t.table_id===diagnostic.scope_id),evidence=entry?.automatic_border_preview
  const table=document.body.blocks.find(b=>b.table?.id===diagnostic.scope_id)?.table
  if(!evidence||!table||table.borders||entry?.borders||evidence.package_sha256!==document.source.package_sha256||evidence.source_part!==resolved.source_parts.styles_part||diagnostic.part_name!==evidence.source_part||!/^\/w:styles\[1\]\/w:style\[[1-9][0-9]*\]\/w:tblPr\[1\]\/w:tblBorders\[1\]$/.test(evidence.source_path)||evidence.source_path!==diagnostic.path+'/w:tblBorders[1]')return false
  const parts=document.passthrough_parts.filter(p=>p.part_name===evidence.source_part),cells=table.rows.flatMap(r=>r.cells)
  if(parts.length!==1||parts[0]!.sha256!==evidence.source_sha256||JSON.stringify(cells.map(c=>c.id))!==JSON.stringify(evidence.cell_ids)||cells.some(c=>c.grid_span!==1||c.vertical_merge!=='none'))return false
  return evidence.source_diagnostics.some(d=>d.code===diagnostic.code&&d.scope_id===diagnostic.scope_id&&d.part_name===diagnostic.part_name&&d.path===diagnostic.path)
 }
 const add=(scope:string,id:string)=>{if(!sources.has(scope)){global.push(id);return}const list=blockers.get(scope)??[];list.push(id);blockers.set(scope,list)}
 for(const diagnostic of document.unsupported){
  if(nestedIDs.has(diagnostic.id))continue // Retained below as exact source-bound omissions.
  const owner=sources.get(diagnostic.scope_id)
  if(owner&&diagnostic.anchor&&(owner.part_name!==diagnostic.anchor.part_name||!diagnostic.anchor.path.startsWith(owner.path+'/')&&diagnostic.anchor.path!==owner.path)){global.push(diagnostic.id);continue}
  add(diagnostic.scope_id,diagnostic.id)
 }
 for(const [index,diagnostic]of (resolved?.diagnostics??[]).entries()){
  const id=`resolved:${index}:${diagnostic.code}`,owner=sources.get(diagnostic.scope_id)
  if(nontext(diagnostic)){retainedNontext.push(id);continue}
  if(owner&&(diagnostic.part_name!==undefined&&diagnostic.part_name!==owner.part_name||diagnostic.path!==undefined&&diagnostic.path!==owner.path&&!diagnostic.path.startsWith(owner.path+'/'))){global.push(id);continue}
  add(diagnostic.scope_id,id)
 }
 const paragraphs=new Map(resolved?.paragraphs.map(p=>[p.paragraph_id,p])),runs=new Map(resolved?.runs.map(r=>[r.run_id,r]))
 const omissions:NativeDocxPartialContentV1['omissions']=[],blocks:NativeDocxPartialContentV1['blocks']=[]
 let textUnits=0,textRuns=0,tableCells=0
 const source=(scope_id:string,anchor:NativeDocxSourceAnchorV1)=>({scope_id,anchor:structuredClone(anchor)})
 const omit=(s:NativeDocxPartialSourceV1,code:NativeDocxPartialOmissionCode,count=1,diagnostic_ids:string[]=[]):Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>=>{
  const result={kind:'omission' as const,source:s,code,count,diagnostic_ids:[...diagnostic_ids]};omissions.push(result);return result
 }
 const projectParagraph=(paragraph:NativeDocxParagraphV1):NativeDocxPartialParagraphV1|Extract<NativeDocxPartialSegmentV1,{kind:'omission'}>=>{
  const s=source(paragraph.id,paragraph.anchor),diagnostics=blockers.get(paragraph.id)??[]
  if(diagnostics.length)return omit(s,'unsupported-source',1,diagnostics)
  const segments:NativeDocxPartialSegmentV1[]=[],p=paragraphs.get(paragraph.id)
  for(const run of paragraph.runs){
   const rs=source(run.id,run.anchor),r=runs.get(run.id),diagnostics=[...(blockers.get(run.id)??[]),...(run.drawing?blockers.get(run.drawing.id)??[]:[])]
   if(diagnostics.length){segments.push(omit(rs,'unsupported-source',1,diagnostics));continue}
   if(run.properties?.hidden||r?.properties.hidden){segments.push(omit(rs,'hidden-text'));continue}
   if(run.page_field){segments.push(omit(rs,'field'));continue}
   if(run.kind!=='text'){
    segments.push(omit(rs,run.kind==='drawing'?'drawing':run.kind==='reference'?'reference':'control'))
    const alt=run.drawing?.alt_text
    if(alt&&p&&r&&r.paragraph_id===paragraph.id){
     const ds=source(run.drawing!.id,run.drawing!.anchor)
     if(textUnits+alt.length>DOCX_PARTIAL_CONTENT_LIMITS.textUnits)segments.push(omit(ds,'text-limit'))
     else{textUnits+=alt.length;segments.push({kind:'alternative-text',source:ds,text:alt,label:'authored-drawing-description'})}
    }
    continue
   }
   if(!p||!r||r.paragraph_id!==paragraph.id){segments.push(omit(rs,'unqualified-text-visibility'));continue}
   const text=run.text??''
   if(textUnits+text.length>DOCX_PARTIAL_CONTENT_LIMITS.textUnits){segments.push(omit(rs,'text-limit'));continue}
   textUnits+=text.length;textRuns++;segments.push({kind:'text',source:rs,text})
  }
  return {kind:'paragraph',source:s,segments}
 }
 const inherited=[...global,...(blockers.get(document.body.id)??[])]
 // Retain global diagnostics once rather than repeating them across blocks.
 if(inherited.length)blocks.push(omit(source(document.body.id,document.body.anchor),'unsupported-source',document.body.blocks.length,inherited))
 for(const block of inherited.length?[]:document.body.blocks.slice(0,DOCX_PARTIAL_CONTENT_LIMITS.bodyBlocks)){
  if(block.paragraph){blocks.push(projectParagraph(block.paragraph));continue}
  const table=block.table!
  const s=source(table.id,table.anchor),diagnostics=blockers.get(table.id)??[]
  if(diagnostics.length){blocks.push(omit(s,'unsupported-source',1,diagnostics));continue}
  const cells:NativeDocxPartialCellV1[]=[],count=table.rows.reduce((n,r)=>n+r.cells.length,0)
  let visited=0
  for(const [rowOrdinal,row]of table.rows.entries())for(const [cellOrdinal,cell]of row.cells.entries()){
   if(tableCells>=DOCX_PARTIAL_CONTENT_LIMITS.tableCells)continue
   tableCells++;visited++
   const cs=source(cell.id,cell.anchor),diagnostics=[...(blockers.get(row.id)??[]),...(blockers.get(cell.id)??[])]
   const entries=[...cell.paragraphs.map(p=>({start:p.anchor.start_byte,paragraph:p})),...nested.items.filter(n=>n.cell_id===cell.id).map(n=>({start:n.anchor.start_byte,nested:n}))].sort((a,b)=>a.start-b.start)
   // Recover only source-owner text, not shared-cell geometry or continuation
   // text. The existing paragraph/run visibility and diagnostic gates still run.
   const content=diagnostics.length?[omit(cs,'unsupported-source',1,diagnostics)]:cell.vertical_merge==='continue'?[omit(cs,'merged-cell')]:entries.map(entry=>'paragraph'in entry?projectParagraph(entry.paragraph):omit(source(entry.nested.diagnostic_id,entry.nested.anchor),'nested-table',1,[entry.nested.diagnostic_id]))
   cells.push({kind:'cell-source',source:cs,row_ordinal:rowOrdinal,cell_ordinal:cellOrdinal,...(cell.grid_span!==1||cell.vertical_merge!=='none'?{source_merge:{grid_span:cell.grid_span,vertical_merge:cell.vertical_merge}}:{}),paragraphs:content})
  }
  blocks.push({kind:'table-source',source:s,source_cell_count:count,cells})
  // Table layout remains explicitly omitted even when cell text is recovered.
  blocks.push(omit(s,'table'))
  if(visited<count)blocks.push(omit(s,'cell-limit',count-visited))
 }
 const remaining=document.body.blocks.length-DOCX_PARTIAL_CONTENT_LIMITS.bodyBlocks
 if(remaining>0&&!inherited.length)blocks.push(omit(source(document.body.id,document.body.anchor),'block-limit',remaining))
 if(nested.omitted_count)blocks.push(omit(source(document.body.id,document.body.anchor),'nested-table-limit',nested.omitted_count))
 for(const story of [...document.headers,...document.footers,...document.notes,...document.comment_stories])blocks.push(omit(source(story.id,story.anchor),'nonbody-story',1,blockers.get(story.id)??[]))
 return {protocol:'injoffice.docx.partial-content',version:1,policy:DOCX_PARTIAL_CONTENT_POLICY,read_only:true,fidelity:'partial-source-content',pagination:'not-produced',source:{document_id:document.document_id,revision:document.revision,package_sha256:document.source.package_sha256},blocks,omissions,source_diagnostics:{document:document.unsupported,resolved:resolved?.diagnostics??[]},retained_nontext_diagnostic_ids:retainedNontext,...(nestedEvidence===undefined?{}:{nested_table_omissions:nested}),coverage:{body_blocks:document.body.blocks.length,visited_body_blocks:inherited.length?0:Math.min(document.body.blocks.length,DOCX_PARTIAL_CONTENT_LIMITS.bodyBlocks),projected_text_runs:textRuns,omitted_units:omissions.reduce((n,o)=>n+o.count,0),layout_present:resolved!==undefined},warnings:['Read-only extracted source content, not document pagination. Formatting, list markers, layout and nonbody stories are not reconstructed. Every omitted unit has a source-bound placeholder.',...(retainedNontext.length?['Source-qualified nontext metadata diagnostics are retained but do not prevent plain text recovery. This does not qualify their rendering.']:[]),...(resolved?[]:['No resolved layout was supplied: text visibility is unqualified, so this result contains an omission inventory rather than text.'])]}
}
