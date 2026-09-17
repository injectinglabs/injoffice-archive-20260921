import {decodeNativeDocxDocument,type NativeDocxSourceAnchorV1} from './nativeContract.js'
import {decodeNativeDocxResolvedLayout} from './nativeResolvedLayout.js'
export interface NativeDocxPartialTableTextContextV1 {package_sha256:string;table_id:string;look_diagnostic_id?:string;look_anchor:NativeDocxSourceAnchorV1;styles_part:string;styles_sha256:string;style_chain:Array<{style_id:string;anchor:NativeDocxSourceAnchorV1}>;resolved_diagnostics:Array<{code:'TABLE_STYLE_EFFECTS_PRESERVED';scope_id:string;part_name:string;path:string}>}
function record(value:unknown,keys:string[],optional:string[]=[]):Record<string,unknown>{
 if(!value||typeof value!=='object'||Object.getPrototypeOf(value)!==Object.prototype||Reflect.ownKeys(value).some(k=>typeof k!=='string'))throw new TypeError('Invalid table text context record')
 const present=Object.keys(value).sort(),allowed=[...keys,...optional]
 if(present.some(k=>!allowed.includes(k))||keys.some(k=>!present.includes(k)))throw new TypeError('Invalid table text context record')
 const out:Record<string,unknown>={};for(const key of present){const d=Object.getOwnPropertyDescriptor(value,key);if(!d||!('value'in d))throw new TypeError('Table context accessors are forbidden');out[key]=d.value}return out
}
function array(value:unknown,max:number):unknown[]{
 if(!Array.isArray(value)||Object.getPrototypeOf(value)!==Array.prototype||value.length>max||Reflect.ownKeys(value).length!==value.length+1)throw new TypeError('Invalid table context array')
 return Array.from({length:value.length},(_,i)=>{const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d||!('value'in d))throw new TypeError('Table context arrays must be dense plain data');return d.value})
}
const anchorKeys=['part_name','path','start_byte','end_byte','xml_sha256']
/** Source-qualified geometry-only style evidence from the same-byte inspector.
 * It authorizes plain-text recovery, never table rendering or mutation. Style
 * slice hashes are producer evidence, joined to the retained source part hash. */
export function decodeNativeDocxPartialTableTextContextsV1(source:unknown,resolvedValue:unknown,input:unknown):NativeDocxPartialTableTextContextV1[]{
 const document=decodeNativeDocxDocument(source);if(!document.ok)throw new TypeError('Invalid table text source')
 if(input===undefined)return []
 const layout=decodeNativeDocxResolvedLayout(resolvedValue),d=document.value
 if(!layout.ok||layout.value.document_id!==d.document_id||layout.value.revision!==d.revision||layout.value.source_parts.main_part!==d.source.main_part)throw new TypeError('Table text context requires exact resolved identity')
 const result:NativeDocxPartialTableTextContextV1[]=[],seen=new Set<string>()
 for(const raw of array(input,64)){
  const v=record(raw,['package_sha256','table_id','look_anchor','styles_part','styles_sha256','style_chain','resolved_diagnostics'],['look_diagnostic_id']),look=record(v.look_anchor,anchorKeys)
  const diagnosed='look_diagnostic_id'in v
  if(v.package_sha256!==d.source.package_sha256||typeof v.table_id!=='string'||seen.has(v.table_id)||diagnosed&&typeof v.look_diagnostic_id!=='string'||v.styles_part!==layout.value.source_parts.styles_part)throw new TypeError('Invalid table text source identity')
  const table=d.body.blocks.find(b=>b.table?.id===v.table_id)?.table,parts=d.passthrough_parts.filter(p=>p.part_name===v.styles_part),part=parts[0]
  if(!table?.table_style_id||parts.length!==1||!part||part.sha256!==v.styles_sha256||look.part_name!==d.source.main_part||look.path!==table.anchor.path+'/w:tblPr[1]/w:tblLook[1]'||typeof look.xml_sha256!=='string'||!/^sha256:[0-9a-f]{64}$/.test(look.xml_sha256)||!Number.isSafeInteger(look.start_byte)||!Number.isSafeInteger(look.end_byte)||Number(look.end_byte)<=Number(look.start_byte)||Number(look.start_byte)<table.anchor.start_byte||Number(look.end_byte)>table.anchor.end_byte)throw new TypeError('Table text context does not join exact source look')
  // The producer omits the diagnostic id exactly when the extractor proved the
  // look inactive, which leaves no diagnostic on that element to name. Either
  // state is source-bound evidence about the same bytes and neither authorizes
  // mutation, so the undiagnosed form is admitted only when the document really
  // carries no unmodeled-property diagnostic for this table's look.
  const anchored=d.unsupported.filter(x=>x.code==='UNMODELED_TABLE_PROPERTY'&&x.scope_id===table.id&&x.anchor!==undefined&&x.anchor.part_name===look.part_name&&x.anchor.path===look.path)
  if(!diagnosed){
   if(anchored.length!==0)throw new TypeError('Table text context omits a look diagnostic the source states')
  }else{
   const diagnostics=d.unsupported.filter(x=>x.id===v.look_diagnostic_id),diag=diagnostics[0]
   if(diagnostics.length!==1||!diag?.anchor||diag.scope_id!==table.id||diag.code!=='UNMODELED_TABLE_PROPERTY'||diag.capability!=='table-properties'||diag.preservation!=='refuse-mutation'||Object.entries(diag.anchor).some(([k,x])=>look[k]!==x))throw new TypeError('Table text context does not join exact source look')
  }
  const chain:NativeDocxPartialTableTextContextV1['style_chain']=[],styleIDs=new Set<string>(),paths=new Set<string>()
  for(const rawStyle of array(v.style_chain,16)){
   const style=record(rawStyle,['style_id','anchor']),a=record(style.anchor,anchorKeys)
   if(typeof style.style_id!=='string'||style.style_id.length===0||style.style_id.length>256||styleIDs.has(style.style_id)||a.part_name!==v.styles_part||typeof a.path!=='string'||!/^\/w:styles\[1\]\/w:style\[[1-9][0-9]*\]$/.test(a.path)||paths.has(a.path)||!Number.isSafeInteger(a.start_byte)||!Number.isSafeInteger(a.end_byte)||Number(a.start_byte)<0||Number(a.end_byte)<=Number(a.start_byte)||Number(a.end_byte)>part.byte_length||typeof a.xml_sha256!=='string'||!/^sha256:[0-9a-f]{64}$/.test(a.xml_sha256))throw new TypeError('Invalid geometry-only style chain source')
   chain.push({style_id:style.style_id,anchor:a as unknown as NativeDocxSourceAnchorV1});styleIDs.add(style.style_id);paths.add(a.path)
  }
  if(chain.length===0||chain.at(-1)!.style_id!==table.table_style_id)throw new TypeError('Table style chain does not join authored selection')
  const resolved:NativeDocxPartialTableTextContextV1['resolved_diagnostics']=[],resolvedPaths=new Set<string>()
  for(const rawDiagnostic of array(v.resolved_diagnostics,16)){
   const r=record(rawDiagnostic,['code','scope_id','part_name','path'])
   if(r.code!=='TABLE_STYLE_EFFECTS_PRESERVED'||r.scope_id!==table.id||r.part_name!==v.styles_part||typeof r.path!=='string'||!chain.some(s=>r.path===s.anchor.path+'/w:tblPr[1]')||resolvedPaths.has(r.path))throw new TypeError('Invalid table geometry diagnostic context')
   const matches=layout.value.diagnostics.filter(x=>x.code===r.code&&x.scope_id===r.scope_id&&x.part_name===r.part_name&&x.path===r.path&&x.severity==='unsupported'&&x.preservation==='preserve-verbatim')
   if(matches.length!==1)throw new TypeError('Table context does not join original resolved diagnostic')
   resolvedPaths.add(r.path);resolved.push(r as NativeDocxPartialTableTextContextV1['resolved_diagnostics'][number])
  }
  result.push({...v,look_anchor:look,style_chain:chain,resolved_diagnostics:resolved} as unknown as NativeDocxPartialTableTextContextV1);seen.add(table.id)
 }
 return result
}
