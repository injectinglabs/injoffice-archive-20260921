import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
import type {NativeDocxShapedLinesV1} from './nativeShapingLines.js'
import {qualifyNativeDocxTablesV1,nativeDocxTableProjectionSha256V1} from './nativeTablePagePaintV1.js'
export const DOCX_LEGACY_TABLE_ORIGIN_WARNING='Approximate read-only preview: eligible legacy tables align their leading cell content to the source indent by shifting the table left by its explicit cell margin; this is not Word-validated layout.'
export const DOCX_TABLE_BORDER_RESERVATION_WARNING='Approximate read-only preview: eligible legacy tables reserve one authored horizontal border width above each row’s content; this is a declared collapsed-border layout policy, not Word-validated layout.'
export interface NativeDocxLegacyTableOriginV1 {table_id:string;package_sha256:string;indent_twips:number;left_margin_twips:number;source_indent:{part_name:string;path:string;sha256:string};source_margin:{part_name:string;path:string;sha256:string}}
const keys=(v:object,names:string)=>Object.keys(v).sort().join(',')===names
export function validLegacyTableOrigins(value:unknown,hash:string):value is NativeDocxLegacyTableOriginV1[]{
 if(!Array.isArray(value)||value.length>1000)return false
 const ids=new Set<string>()
 return value.every(f=>{
  if(!f||typeof f!=='object'||!keys(f,'indent_twips,left_margin_twips,package_sha256,source_indent,source_margin,table_id')||typeof f.table_id!=='string'||f.table_id.length>256||ids.has(f.table_id)||f.package_sha256!==hash||!Number.isSafeInteger(f.indent_twips)||f.indent_twips<0||f.indent_twips>20000000||!Number.isSafeInteger(f.left_margin_twips)||f.left_margin_twips<1||f.left_margin_twips>20000000)return false
  ids.add(f.table_id)
  return [f.source_indent,f.source_margin].every(s=>s&&typeof s==='object'&&keys(s,'part_name,path,sha256')&&typeof s.part_name==='string'&&/^[A-Za-z0-9_.\/-]+$/.test(s.part_name)&&!s.part_name.split('/').includes('..')&&typeof s.path==='string'&&s.path.length<4096&&s.path.startsWith('/')&&typeof s.sha256==='string'&&/^sha256:[0-9a-f]{64}$/.test(s.sha256))
 })
}

/** Internal explicit approximation. The public strict qualifier remains unchanged. */
export function qualifyApproximateLegacyTables(document:NativeDocxDocumentV1,resolved:NativeDocxResolvedLayoutInputV1,shaped:NativeDocxShapedLinesV1|undefined,eligibility?:{legacy_compatibility_mode:number|null;legacy_table_origins?:NativeDocxLegacyTableOriginV1[]}){
 const result=qualifyNativeDocxTablesV1(document,resolved,shaped),facts=eligibility?.legacy_table_origins??[]
 if(result.status!=='qualified'||eligibility?.legacy_compatibility_mode!==12&&!facts.length)return result
 if(eligibility?.legacy_compatibility_mode!==12||!validLegacyTableOrigins(facts,document.source.package_sha256))throw new TypeError('Invalid legacy table origin eligibility')
 const tables=structuredClone(result.tables),sections=new Map(document.sections.map(s=>[s.starts_at_block_id,s]))
 const shapedParagraphs=new Map(shaped?.paragraphs.map(paragraph=>[paragraph.paragraph_id,paragraph]))
 let section:NativeDocxDocumentV1['sections'][number]|undefined
 const owners=new Map<string,NativeDocxDocumentV1['sections'][number]>()
 for(const block of document.body.blocks){section=sections.get(block.id)??section;if(block.table&&section)owners.set(block.table.id,section)}
 for(const entry of tables){
  const owner=owners.get(entry.table.id)
  if(!owner||owner.page.columns!==1||owner.page.margins.gutter_twips!==0)continue
  const borders=entry.table.borders,top=borders?.top,inside=borders?.inside_horizontal,bottom=borders?.bottom
  // The strict qualifier already excludes floating/spacing/unknown geometry.
  // Keep the first policy bounded to identical, explicit single horizontal
  // borders and automatic unmerged rows; do not infer conflict resolution.
  if(top?.style==='single'&&inside?.style==='single'&&bottom?.style==='single'&&top.size_eighth_points===inside.size_eighth_points&&top.size_eighth_points===bottom.size_eighth_points&&entry.table.rows.every(row=>row.height_rule===undefined&&row.height_twips===undefined&&row.cells.every(cell=>cell.grid_span===1&&cell.vertical_merge==='none'&&cell.paragraphs.length===1&&shapedParagraphs.get(cell.paragraphs[0]!.id)?.lines.length===1))){
   entry.border_reservation_policy={name:'collapsed-horizontal-border-reservation-v1',above_content_millipoints:top.size_eighth_points*125}
  }
 }
 for(const fact of facts){
  const entry=tables.find(t=>t.table.id===fact.table_id),owner=owners.get(fact.table_id)
  if(!entry||!owner||owner.page.columns!==1||owner.page.margins.gutter_twips!==0||entry.table.alignment!=='left'||entry.table.indent_twips!==fact.indent_twips||entry.table.cell_margins?.left_twips!==fact.left_margin_twips)throw new TypeError('Legacy origin does not join qualified source geometry')
  if(entry.table.rows.some(row=>row.cells.some(cell=>cell.grid_span!==1||cell.vertical_merge!=='none')))throw new TypeError('Legacy origin requires unmerged cells')
  for(const [source,leaf]of [[fact.source_indent,'/w:tblInd[1]'],[fact.source_margin,'/w:tblCellMar[1]/w:left[1]']] as const){
   if(!source.path.endsWith(leaf)||![document.source.main_part,resolved.source_parts.styles_part].includes(source.part_name))throw new TypeError('Legacy origin source owner mismatch')
   const direct=source.part_name===document.source.main_part,expected=entry.table.anchor.path+'/w:tblPr[1]'+leaf
   if(direct?source.path!==expected:!/^\/w:styles\[1\]\/w:style\[[1-9][0-9]*\]\/w:tblPr\[1\]$/.test(source.path.slice(0,-leaf.length)))throw new TypeError('Legacy origin source path mismatch')
   const parts=document.passthrough_parts.filter(p=>p.part_name===source.part_name)
   if((!direct||parts.length>0)&&(parts.length!==1||parts[0]!.sha256!==source.sha256))throw new TypeError('Legacy origin source digest mismatch')
   // A modeled main part has no independent whole-part hash in this input.
   // Its digest is trusted extractor evidence bound to the exact package,
   // not a digest verified against XML bytes by this consumer.
  }
  const delta=-fact.left_margin_twips*50,x=entry.x_millipoints+delta,pageLeft=owner.page.margins.left_twips*50
  if(pageLeft+x<0||pageLeft+x+entry.width_millipoints>owner.page.width_twips*50)throw new TypeError('Legacy origin exceeds page bounds')
  entry.origin_policy={name:'legacy-content-aligned-origin-v1',source:structuredClone(fact),delta_millipoints:delta}
  entry.x_millipoints=x
  for(const row of entry.rows)for(const cell of row.cells){cell.x_millipoints+=delta;cell.content_x_millipoints+=delta}
 }
 return {...result,tables,sha256:nativeDocxTableProjectionSha256V1(tables)}
}
