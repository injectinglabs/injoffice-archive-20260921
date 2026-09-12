import {snapshotNativePlainData} from './nativePlainData.js'
import {EXCEL_MAX_ROWS,EXCEL_MAX_COLUMNS} from './mutationProtocol.js'

export interface NativeSheetPrintTitleRangeV1 {start:number;end:number}
export type NativeSheetPrintTitlesV1 = {
 sheet_id:string;sheet_part:string;warnings:string[];
} & ({status:'available';rows?:NativeSheetPrintTitleRangeV1;columns?:NativeSheetPrintTitleRangeV1}|{status:'unavailable';rows?:never;columns?:never})

/** Snapshot only bounded worksheet-local title metadata; package identity comes from its envelope. */
export function decodeNativeSheetPrintTitlesV1(input:unknown):NativeSheetPrintTitlesV1[]{
 const value=snapshotNativePlainData(input,{maxDepth:6,maxNodes:4096})
 const fail=():never=>{throw new TypeError('Invalid native worksheet print titles')}
 const exact=(v:unknown,keys:string[])=>{
  if(!v||typeof v!=='object'||Array.isArray(v))return fail()
  const o=v as Record<string,unknown>
  if(Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k)))return fail()
  return o
 }
 if(!Array.isArray(value)||value.length>64)return fail()
 const ids=new Set<string>(),parts=new Set<string>()
 return value.map(v=>{
  const raw=v as Record<string,unknown>,status=raw?.status
  const axes=['rows','columns'].filter(k=>raw&&Object.hasOwn(raw,k))
  const o=exact(v,['sheet_id','sheet_part','status','warnings',...(status==='available'?axes:[])])
  if(typeof o.sheet_id!=='string'||!/^[1-9][0-9]{0,9}$/.test(o.sheet_id)||Number(o.sheet_id)>0xffffffff||ids.has(o.sheet_id)||typeof o.sheet_part!=='string'||o.sheet_part.length>1024||!/^[A-Za-z0-9_.\/-]+$/.test(o.sheet_part)||o.sheet_part.split('/').some(s=>!s||s==='.'||s==='..')||parts.has(o.sheet_part)||(status!=='available'&&status!=='unavailable'))return fail()
  ids.add(o.sheet_id);parts.add(o.sheet_part)
  if(!Array.isArray(o.warnings)||o.warnings.length<1||o.warnings.length>8||o.warnings.some(w=>typeof w!=='string'||w.length>4096||/[\u0000-\u001f\u007f]/.test(w)))return fail()
  const base={sheet_id:o.sheet_id,sheet_part:o.sheet_part,warnings:o.warnings as string[]}
  if(status==='unavailable')return {...base,status}
  if(!axes.length)return fail()
  const ranges:Pick<NativeSheetPrintTitlesV1,'rows'|'columns'>={}
  for(const axis of axes){
   const r=exact(o[axis],['start','end']),limit=axis==='rows'?EXCEL_MAX_ROWS:EXCEL_MAX_COLUMNS
   for(const n of [r.start,r.end])if(typeof n!=='number'||!Number.isSafeInteger(n)||Object.is(n,-0)||n<0||n>=limit)return fail()
   if(Number(r.start)>Number(r.end))return fail()
   ranges[axis as 'rows'|'columns']={start:Number(r.start),end:Number(r.end)}
  }
  return {...base,status,...ranges}
 })
}
