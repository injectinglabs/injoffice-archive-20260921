import {decodeNativeDocxDocument,type NativeDocxSourceAnchorV1} from './nativeContract.js'
import {decodeNativeDocxResolvedLayout} from './nativeResolvedLayout.js'
import {isRenderNeutralLayoutDiagnostic} from './nativeRenderDiagnostics.js'

export type NativeDocxMathNodeV1={kind:'text';text:string}|{kind:'row'|'fraction'|'superscript'|'subscript'|'radical';children:NativeDocxMathNodeV1[]}
export interface NativeDocxEquationPreviewV1 {package_sha256:string;paragraph_id:string;anchor:NativeDocxSourceAnchorV1;diagnostic_id:string;status:'supported'|'omitted';tree?:NativeDocxMathNodeV1;reason?:string;context_notice_ids?:string[]}
export interface NativeDocxEquationContextNoticeV1 {kind:'horizontal-section'|'ignored-font-matching'|'disabled-paragraph-hyphenation'|'unused-paragraph-tab-stops';package_sha256:string;part_sha256:string;anchor:NativeDocxSourceAnchorV1;diagnostic_origin:'document'|'resolved';diagnostic_id?:string;code:string;scope_id:string;value?:string;character_set?:string;tab_stops?:Array<{kind:string;position_twips:number;leader:'none'}>}
const noticeID=(n:NativeDocxEquationContextNoticeV1)=>n.diagnostic_origin==='document'?`document:${n.diagnostic_id}`:`resolved:${n.code}:${n.scope_id}:${n.anchor.part_name}:${n.anchor.path}`
function own(value:unknown):value is Record<string,unknown>{
 if(!value||typeof value!=='object'||Object.getPrototypeOf(value)!==Object.prototype)return false
 return Reflect.ownKeys(value).every(k=>typeof k==='string'&&'value'in Object.getOwnPropertyDescriptor(value,k)!)
}
function exact(value:Record<string,unknown>,keys:string[]):boolean{return Object.keys(value).sort().join(',')===keys.sort().join(',')}
function snapshot(value:unknown,depth=0,budget={nodes:0}):unknown{
 if(depth>80||++budget.nodes>100_000)throw new TypeError('Equation evidence exceeds structural budget')
 if(value===null||typeof value==='string'||typeof value==='number'||typeof value==='boolean')return value
 if(Array.isArray(value)){
  if(Object.getPrototypeOf(value)!==Array.prototype||value.length>1000||Reflect.ownKeys(value).length!==value.length+1)throw new TypeError('Equation arrays must be dense plain data')
  const result:unknown[]=[]
  for(let i=0;i<value.length;i++){const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d||!('value'in d))throw new TypeError('Equation array accessors are forbidden');result.push(snapshot(d.value,depth+1,budget))}
  return result
 }
 if(!own(value))throw new TypeError('Equation evidence must contain own plain data')
 const result:Record<string,unknown>={}
 for(const key of Object.keys(value)){if(key==='__proto__')throw new TypeError('Invalid equation key');result[key]=snapshot(Object.getOwnPropertyDescriptor(value,key)!.value,depth+1,budget)}
 return result
}
function validateTree(value:unknown,depth:number,budget:{nodes:number;units:number}):value is NativeDocxMathNodeV1 {
 if(depth>32||++budget.nodes>256||!own(value))return false
 if(value.kind==='text'){if(!exact(value,['kind','text'])||typeof value.text!=='string'||!value.text.length)return false;budget.units+=new TextEncoder().encode(value.text).length;return budget.units<=32768}
 if(!exact(value,['kind','children'])||!Array.isArray(value.children)||!value.children.length)return false
 const arity=value.kind==='row'?undefined:value.kind==='radical'?1:['fraction','superscript','subscript'].includes(String(value.kind))?2:0
 if(arity===0||arity!==undefined&&value.children.length!==arity)return false
 return value.children.length<=256&&value.children.every(v=>validateTree(v,depth+1,budget))
}

/** Read-only, unpaginated browser math approximation. No authored markup or
 * attributes are returned as executable HTML; strict source diagnostics stay. */
export function decodeNativeDocxEquationContextNoticesV1(source:unknown,layout:unknown,input:unknown):NativeDocxEquationContextNoticeV1[]{
 const d=decodeNativeDocxDocument(source),r=decodeNativeDocxResolvedLayout(layout)
 if(!d.ok||!r.ok||d.value.document_id!==r.value.document_id||d.value.revision!==r.value.revision||d.value.source.main_part!==r.value.source_parts.main_part)throw new TypeError('Context notices require joined native source')
 if(input===undefined)return []
 input=snapshot(input);if(!Array.isArray(input)||input.length>1000)throw new TypeError('Invalid context notice budget')
 const seen=new Set<string>(),hash=/^sha256:[0-9a-f]{64}$/
 return input.map(v=>{
  if(!own(v)||!own(v.anchor))throw new TypeError('Invalid context notice')
  const extra=v.kind==='unused-paragraph-tab-stops'?['tab_stops']:v.kind==='ignored-font-matching'?['value','character_set']:['value']
  if(!exact(v,['kind','package_sha256','part_sha256','anchor','diagnostic_origin','code','scope_id',...(v.diagnostic_origin==='document'?['diagnostic_id']:[]),...extra])||v.package_sha256!==d.value.source.package_sha256||!hash.test(String(v.part_sha256))||typeof v.code!=='string'||typeof v.scope_id!=='string')throw new TypeError('Invalid source-bound context notice')
  const a=v.anchor
  if(!exact(a,['part_name','path','start_byte','end_byte','xml_sha256'])||typeof a.part_name!=='string'||typeof a.path!=='string'||!hash.test(String(a.xml_sha256))||!Number.isSafeInteger(a.start_byte)||!Number.isSafeInteger(a.end_byte)||Number(a.start_byte)<0||Number(a.end_byte)<=Number(a.start_byte))throw new TypeError('Invalid context anchor')
  if(v.diagnostic_origin==='document'){
   const diag=d.value.unsupported.find(x=>x.id===v.diagnostic_id),section=d.value.sections.find(s=>s.id===v.scope_id)
   if(!diag?.anchor||diag.code!==v.code||diag.scope_id!==v.scope_id||diag.preservation!=='refuse-mutation'||Object.entries(diag.anchor).some(([k,x])=>a[k]!==x)||v.kind!=='horizontal-section'||v.code!=='UNMODELED_SECTION_PROPERTY'||v.value!=='lrTb'||a.part_name!==d.value.source.main_part||!section||a.path!==section.anchor.path+'/w:textDirection[1]')throw new TypeError('Unqualified horizontal section notice')
   // Modeled main-part digest is producer evidence, not independently rehashed
   // here. The original diagnostic slice digest and caller package join bind it.
  }else if(v.diagnostic_origin==='resolved'){
   const matches=r.value.diagnostics.filter(x=>x.code===v.code&&x.scope_id===v.scope_id&&x.part_name===a.part_name&&x.path===a.path&&x.severity==='unsupported'&&x.preservation==='preserve-verbatim'),parts=d.value.passthrough_parts.filter(p=>p.part_name===a.part_name)
   if(matches.length!==1||parts.length!==1||parts[0]!.sha256!==v.part_sha256||Number(a.end_byte)>parts[0]!.byte_length)throw new TypeError('Context diagnostic/part digest does not join')
   if(v.kind==='ignored-font-matching'){
    if(v.code!=='UNMODELED_FONT_METADATA'||v.scope_id!==d.value.document_id||a.part_name!==r.value.source_parts.font_table_part||!/^\/w:fonts\[1\]\/w:font\[[1-9][0-9]*\]\/w:charset\[1\]$/.test(a.path)||!(v.value==='00'&&v.character_set==='windows-1252'||v.value==='80'&&v.character_set==='utf-8'))throw new TypeError('Unqualified ignored font matching notice')
   }else{
    const leaf=v.kind==='disabled-paragraph-hyphenation'?'suppressAutoHyphens':v.kind==='unused-paragraph-tab-stops'?'tabs':''
    if(!leaf||v.code!=='UNMODELED_PARAGRAPH_PROPERTY'||!r.value.paragraphs.some(p=>p.paragraph_id===v.scope_id)||a.part_name!==r.value.source_parts.styles_part||!new RegExp('^/w:styles\\[1\\]/w:style\\[[1-9][0-9]*\\]/w:pPr\\[1\\]/w:'+leaf+'\\[1\\]$').test(a.path))throw new TypeError('Unqualified paragraph context notice')
    if(leaf==='suppressAutoHyphens'&&v.value!=='true')throw new TypeError('Hyphenation is not explicitly disabled')
    if(leaf==='tabs'){
     if(!Array.isArray(v.tab_stops)||!v.tab_stops.length||v.tab_stops.length>64)throw new TypeError('Invalid tab context budget')
     const positions=new Set<number>();for(const t of v.tab_stops){if(!own(t)||!exact(t,['kind','position_twips','leader'])||!['left','right','center','decimal','clear'].includes(String(t.kind))||t.leader!=='none'||!Number.isSafeInteger(t.position_twips)||Number(t.position_twips)<0||Number(t.position_twips)>180143985094819||positions.has(Number(t.position_twips)))throw new TypeError('Unqualified tab context');positions.add(Number(t.position_twips))}
    }
   }
  }else throw new TypeError('Invalid context diagnostic origin')
  const result=v as unknown as NativeDocxEquationContextNoticeV1,id=noticeID(result);if(seen.has(id))throw new TypeError('Duplicate context notice');seen.add(id);return result
 })
}
export function createNativeDocxEquationPreviewsV1(source:unknown,layout:unknown,equations:unknown,contextNotices?:unknown):NativeDocxEquationPreviewV1[]{
 const d=decodeNativeDocxDocument(source),r=decodeNativeDocxResolvedLayout(layout)
 if(!d.ok||!r.ok||d.value.document_id!==r.value.document_id||d.value.revision!==r.value.revision||d.value.source.main_part!==r.value.source_parts.main_part)throw new TypeError('Equation preview requires joined native source')
 if(equations===undefined)return []
 equations=snapshot(equations)
 if(!Array.isArray(equations)||equations.length>1000)throw new TypeError('Equation evidence exceeds budget')
 const document=d.value,resolved=r.value,seen=new Set<string>()
 const notices=decodeNativeDocxEquationContextNoticesV1(document,resolved,contextNotices)
 const contextIDs=new Set(notices.map(noticeID))
 const paragraphs=document.body.blocks.flatMap(b=>b.paragraph?[b.paragraph]:b.table!.rows.flatMap(row=>row.cells.flatMap(c=>c.paragraphs)))
 let supported=0
 const validated=equations.map(value=>{
  if(!own(value)||!exact(value,['package_sha256','paragraph_id','anchor','diagnostic_id','status',...(value.status==='supported'?['tree']:['reason'])])||value.package_sha256!==document.source.package_sha256||typeof value.paragraph_id!=='string'||typeof value.diagnostic_id!=='string'||seen.has(value.diagnostic_id))throw new TypeError('Invalid equation source fact')
  seen.add(value.diagnostic_id)
  const p=paragraphs.find(p=>p.id===value.paragraph_id),diagnostic=document.unsupported.find(x=>x.id===value.diagnostic_id)
  const anchor=value.anchor
  if(!p||!own(anchor)||!exact(anchor,['part_name','path','start_byte','end_byte','xml_sha256'])||!diagnostic?.anchor||diagnostic.code!=='UNMODELED_PARAGRAPH_CONTENT'||diagnostic.scope_id!==p.id||Object.entries(diagnostic.anchor).some(([k,v])=>anchor[k]!==v)||anchor.part_name!==p.anchor.part_name||typeof anchor.path!=='string'||!anchor.path.startsWith(p.anchor.path+'/')||anchor.path.slice(p.anchor.path.length+1).includes('/')||!/:oMath(?:Para)?\[[1-9][0-9]*\]$/.test(anchor.path)||Number(anchor.start_byte)<p.anchor.start_byte||Number(anchor.end_byte)>p.anchor.end_byte)throw new TypeError('Equation source anchor does not join')
  if(value.status==='supported'){if(++supported>100||!validateTree(value.tree,0,{nodes:0,units:0}))throw new TypeError('Invalid bounded equation tree')}
  else if(value.status!=='omitted'||typeof value.reason!=='string'||!value.reason.length||value.reason.length>256)throw new TypeError('Invalid equation omission')
  return structuredClone(value) as unknown as NativeDocxEquationPreviewV1
 })
 const qualifiedDiagnosticIDs=new Set(validated.map(e=>e.diagnostic_id))
 const equationParagraphIDs=new Set(validated.map(e=>e.paragraph_id))
 if(notices.some(n=>(n.kind==='disabled-paragraph-hyphenation'||n.kind==='unused-paragraph-tab-stops')&&!equationParagraphIDs.has(n.scope_id)))throw new TypeError('Paragraph context notice has no equation consumer')
 return validated.map(result=>{
  const mark=resolved.paragraphs.find(x=>x.paragraph_id===result.paragraph_id)
  // Unknown document/style/table visibility remains blocking. Only exact
  // equation-source siblings and already qualified nontext metadata are inert.
  const unsafe=document.unsupported.some(x=>!qualifiedDiagnosticIDs.has(x.id)&&!contextIDs.has(`document:${x.id}`))||resolved.diagnostics.some(x=>!isRenderNeutralLayoutDiagnostic(x,resolved)&&!contextIDs.has(`resolved:${x.code}:${x.scope_id}:${x.part_name}:${x.path}`))
  if(notices.length)result.context_notice_ids=[...contextIDs]
  if(!mark||mark.paragraph_mark_properties.hidden||unsafe){delete result.tree;result.status='omitted';result.reason='Equation visibility or surrounding source is not qualified'}
  return result
 })
}
