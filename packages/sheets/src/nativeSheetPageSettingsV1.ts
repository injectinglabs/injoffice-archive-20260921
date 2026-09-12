import {snapshotNativePlainData} from './nativePlainData.js'

export interface NativeSheetPageConfigV1 {
 paper:'Letter'|'A4'; orientation:'portrait'|'landscape'; scale:number;
 left_inches:number;right_inches:number;top_inches:number;bottom_inches:number;
 /** Omission retains down-then-over ordering. */
 page_order?:'downThenOver'|'overThenDown';
}
export interface NativeSheetPageSettingsV1 {
 sheet_id:string;sheet_part:string;status:'available'|'unavailable';
 settings?:NativeSheetPageConfigV1;warnings:string[];
}
/** A standalone bounded copy; package identity is joined by its enclosing object response. */
export function decodeNativeSheetPageSettingsV1(input:unknown):NativeSheetPageSettingsV1[]{
 const value=snapshotNativePlainData(input,{maxDepth:6,maxNodes:4096})
 const fail=():never=>{throw new TypeError('Invalid native worksheet page settings')}
 const exact=(v:unknown,keys:string[])=>{
  if(!v||typeof v!=='object'||Array.isArray(v))return fail()
  const o=v as Record<string,unknown>
  if(Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k)))return fail()
  return o
 }
 if(!Array.isArray(value)||value.length>64)return fail()
 const ids=new Set<string>(),parts=new Set<string>()
 return value.map(v=>{
  const status=(v as Record<string,unknown>)?.status
  const o=exact(v,['sheet_id','sheet_part','status','warnings',...(status==='available'?['settings']:[])])
  if(typeof o.sheet_id!=='string'||!/^[1-9][0-9]{0,9}$/.test(o.sheet_id)||Number(o.sheet_id)>0xffffffff||ids.has(o.sheet_id)||typeof o.sheet_part!=='string'||o.sheet_part.length>1024||!/^[A-Za-z0-9_.\/-]+$/.test(o.sheet_part)||o.sheet_part.split('/').some(s=>!s||s==='.'||s==='..')||parts.has(o.sheet_part)||(status!=='available'&&status!=='unavailable'))return fail()
  ids.add(o.sheet_id);parts.add(o.sheet_part)
  if(!Array.isArray(o.warnings)||o.warnings.length<1||o.warnings.length>8||o.warnings.some(w=>typeof w!=='string'||w.length>4096))return fail()
  let settings:NativeSheetPageConfigV1|undefined
  if(status==='available'){
   const hasOrder=!!o.settings&&typeof o.settings==='object'&&Object.hasOwn(o.settings,'page_order')
   const s=exact(o.settings,['paper','orientation','scale','left_inches','right_inches','top_inches','bottom_inches',...(hasOrder?['page_order']:[])])
   if(hasOrder&&s.page_order!=='downThenOver'&&s.page_order!=='overThenDown')return fail()
   if((s.paper!=='Letter'&&s.paper!=='A4')||(s.orientation!=='portrait'&&s.orientation!=='landscape')||!Number.isInteger(s.scale)||Number(s.scale)<10||Number(s.scale)>400)return fail()
   for(const key of ['left_inches','right_inches','top_inches','bottom_inches'])if(typeof s[key]!=='number'||!Number.isFinite(s[key])||Number(s[key])<0||Number(s[key])>20)return fail()
   settings=s as unknown as NativeSheetPageConfigV1
  }
  return {sheet_id:o.sheet_id,sheet_part:o.sheet_part,status,warnings:o.warnings as string[],...(settings?{settings}:{})}
 })
}
