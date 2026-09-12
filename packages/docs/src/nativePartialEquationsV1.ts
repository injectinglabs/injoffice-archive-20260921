import {decodeNativeDocxDocument,type NativeDocxSourceAnchorV1} from './nativeContract.js'
import {decodeNativeDocxResolvedLayout} from './nativeResolvedLayout.js'
import {isRenderNeutralLayoutDiagnostic} from './nativeRenderDiagnostics.js'

export type NativeDocxMathNodeV1={kind:'text';text:string}|{kind:'row'|'fraction'|'superscript'|'subscript'|'radical';children:NativeDocxMathNodeV1[]}
export interface NativeDocxEquationPreviewV1 {package_sha256:string;paragraph_id:string;anchor:NativeDocxSourceAnchorV1;diagnostic_id:string;status:'supported'|'omitted';tree?:NativeDocxMathNodeV1;reason?:string}
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
export function createNativeDocxEquationPreviewsV1(source:unknown,layout:unknown,equations:unknown):NativeDocxEquationPreviewV1[]{
 const d=decodeNativeDocxDocument(source),r=decodeNativeDocxResolvedLayout(layout)
 if(!d.ok||!r.ok||d.value.document_id!==r.value.document_id||d.value.revision!==r.value.revision||d.value.source.main_part!==r.value.source_parts.main_part)throw new TypeError('Equation preview requires joined native source')
 if(equations===undefined)return []
 equations=snapshot(equations)
 if(!Array.isArray(equations)||equations.length>1000)throw new TypeError('Equation evidence exceeds budget')
 const document=d.value,resolved=r.value,seen=new Set<string>()
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
 return validated.map(result=>{
  const mark=resolved.paragraphs.find(x=>x.paragraph_id===result.paragraph_id)
  // Unknown document/style/table visibility remains blocking. Only exact
  // equation-source siblings and already qualified nontext metadata are inert.
  const unsafe=document.unsupported.some(x=>!qualifiedDiagnosticIDs.has(x.id))||resolved.diagnostics.some(x=>!isRenderNeutralLayoutDiagnostic(x,resolved))
  if(!mark||mark.paragraph_mark_properties.hidden||unsafe){delete result.tree;result.status='omitted';result.reason='Equation visibility or surrounding source is not qualified'}
  return result
 })
}
