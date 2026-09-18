import {isNativePreviewPartPathV1} from './nativePreviewPartPathV1.js'
import {snapshotNativePlainData} from './nativePlainData.js'
import type {NativeDrawingMarkerV1} from './nativeDrawingObjectsV1.js'
/** A worksheet form control, joined from its <controls> anchor, its ctrlProps
 * part and the legacy VML drawing that carries its caption. Only 'checkbox'
 * may be painted; 'unsupported' names what the source stated and nothing is
 * drawn for it. */
export interface NativeFormControlV1 {
 sheet_id:string;sheet_part:string;ordinal:number;shape_id:string;name:string
 kind:'checkbox'|'unsupported'
 object_type?:string
 checked:boolean
 anchor?:{kind:'twoCellAnchor';from:NativeDrawingMarkerV1;to:NativeDrawingMarkerV1}
 caption?:string;caption_size_points?:number
 caption_align?:'left'|'center'|'right';caption_valign?:'top'|'center'|'bottom'
 control_part?:string;legacy_part?:string
 warnings:string[]
}
export function decodeNativeFormControlsV1(input:unknown):NativeFormControlV1[]{
 const value=snapshotNativePlainData(input,{maxDepth:8,maxNodes:15000})
 const fail=():never=>{throw new TypeError('Invalid native source form controls')}
 const exact=(v:unknown,keys:string[])=>{if(!v||typeof v!=='object'||Array.isArray(v))return fail();const o=v as Record<string,unknown>;if(Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k)))return fail();return o}
 const int=(v:unknown,max:number,min=0):number=>typeof v==='number'&&Number.isSafeInteger(v)&&v>=min&&v<=max?v:fail()
 const part=(v:unknown):string=>isNativePreviewPartPathV1(v,false)?v:fail()
 const str=(v:unknown,max:number):string=>typeof v==='string'&&v.length<=max&&!/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(v)?v:fail()
 const marker=(v:unknown):NativeDrawingMarkerV1=>{const m=exact(v,['column','row','column_offset_emu','row_offset_emu']);return {column:int(m.column,16383),row:int(m.row,1048575),column_offset_emu:int(m.column_offset_emu,2147483647),row_offset_emu:int(m.row_offset_emu,2147483647)}}
 if(!Array.isArray(value)||value.length>256)return fail()
 const seen=new Set<string>()
 return value.map(v=>{
  const base=v as Record<string,unknown>
  const has=(key:string)=>!!base&&typeof base==='object'&&Object.hasOwn(base,key)
  const optional=['object_type','anchor','caption','caption_size_points','caption_align','caption_valign','control_part','legacy_part'].filter(has)
  const o=exact(v,['sheet_id','sheet_part','ordinal','shape_id','name','kind','checked','warnings',...optional])
  if(typeof o.sheet_id!=='string'||!/^[1-9][0-9]{0,9}$/.test(o.sheet_id)||Number(o.sheet_id)>0xffffffff)return fail()
  if(o.kind!=='checkbox'&&o.kind!=='unsupported')return fail()
  if(typeof o.checked!=='boolean')return fail()
  if(!Array.isArray(o.warnings)||o.warnings.length<1||o.warnings.length>8||o.warnings.some(w=>typeof w!=='string'||w.length>4096))return fail()
  const sheet_part=part(o.sheet_part),ordinal=int(o.ordinal,256,1)
  const key=JSON.stringify([o.sheet_id,sheet_part,ordinal]);if(seen.has(key))return fail();seen.add(key)
  let anchor:NativeFormControlV1['anchor']
  if(has('anchor')){
   const a=exact(o.anchor,['kind','from','to'])
   if(a.kind!=='twoCellAnchor')return fail()
   anchor={kind:'twoCellAnchor',from:marker(a.from),to:marker(a.to)}
  }
  const align=has('caption_align')?str(o.caption_align,16):undefined
  const valign=has('caption_valign')?str(o.caption_valign,16):undefined
  if(align!==undefined&&align!=='left'&&align!=='center'&&align!=='right')return fail()
  if(valign!==undefined&&valign!=='top'&&valign!=='center'&&valign!=='bottom')return fail()
  const size=has('caption_size_points')?o.caption_size_points:undefined
  if(size!==undefined&&(typeof size!=='number'||!Number.isFinite(size)||size<=0||size>400))return fail()
  const control:NativeFormControlV1={
   sheet_id:o.sheet_id,sheet_part,ordinal,shape_id:str(o.shape_id,32),name:str(o.name,512),
   kind:o.kind,checked:o.checked,warnings:o.warnings as string[],
   ...(has('object_type')?{object_type:str(o.object_type,64)}:{}),
   ...(anchor?{anchor}:{}),
   ...(has('caption')?{caption:str(o.caption,512)}:{}),
   ...(size!==undefined?{caption_size_points:size}:{}),
   ...(align!==undefined?{caption_align:align as 'left'|'center'|'right'}:{}),
   ...(valign!==undefined?{caption_valign:valign as 'top'|'center'|'bottom'}:{}),
   ...(has('control_part')?{control_part:part(o.control_part)}:{}),
   ...(has('legacy_part')?{legacy_part:part(o.legacy_part)}:{}),
  }
  // A painted checkbox must carry every fact the painter reads. Anything
  // short of that is 'unsupported', never a box drawn from a guess.
  if(control.kind==='checkbox'&&(!control.anchor||!control.caption||control.caption_size_points===undefined||control.caption_align===undefined||control.caption_valign===undefined))return fail()
  return control
 })
}
