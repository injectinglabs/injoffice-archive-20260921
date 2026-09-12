import {snapshotNativePlainData} from './nativePlainData.js'
import {EXCEL_MAX_ROWS,EXCEL_MAX_COLUMNS} from './mutationProtocol.js'
import type {NativeSheetViewportV2} from './nativeSheetGeometryV2.js'

export type NativeSheetPrintAreaV1 = {
 sheet_id:string;sheet_part:string;warnings:string[];
} & ({status:'available';area:NativeSheetViewportV2}|{status:'unavailable';area?:never})

/** Bounded plain-data copy. The enclosing object response supplies package identity. */
export function decodeNativeSheetPrintAreasV1(input:unknown):NativeSheetPrintAreaV1[]{
 const value=snapshotNativePlainData(input,{maxDepth:6,maxNodes:4096})
 const fail=():never=>{throw new TypeError('Invalid native worksheet print areas')}
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
  const o=exact(v,['sheet_id','sheet_part','status','warnings',...(status==='available'?['area']:[])])
  if(typeof o.sheet_id!=='string'||!/^[1-9][0-9]{0,9}$/.test(o.sheet_id)||Number(o.sheet_id)>0xffffffff||ids.has(o.sheet_id)||typeof o.sheet_part!=='string'||o.sheet_part.length>1024||!/^[A-Za-z0-9_.\/-]+$/.test(o.sheet_part)||o.sheet_part.split('/').some(s=>!s||s==='.'||s==='..')||parts.has(o.sheet_part)||(status!=='available'&&status!=='unavailable'))return fail()
  ids.add(o.sheet_id);parts.add(o.sheet_part)
  if(!Array.isArray(o.warnings)||o.warnings.length<1||o.warnings.length>8||o.warnings.some(w=>typeof w!=='string'||w.length>4096||/[\u0000-\u001f\u007f]/.test(w)))return fail()
  const base={sheet_id:o.sheet_id,sheet_part:o.sheet_part,warnings:o.warnings as string[]}
  if(status==='unavailable')return {...base,status}
  const a=exact(o.area,['row','column','end_row','end_column'])
  for(const key of ['row','column','end_row','end_column']){
   const n=a[key],limit=key.endsWith('row')?EXCEL_MAX_ROWS:EXCEL_MAX_COLUMNS
   if(typeof n!=='number'||!Number.isSafeInteger(n)||Object.is(n,-0)||n<0||n>=limit)return fail()
  }
  if(Number(a.row)>Number(a.end_row)||Number(a.column)>Number(a.end_column))return fail()
  return {...base,status,area:{row:Number(a.row),column:Number(a.column),end_row:Number(a.end_row),end_column:Number(a.end_column)}}
 })
}
