import {snapshotNativePlainData} from './nativePlainData.js'
import {decodeNativeSheetPrintAreasV1} from './nativeSheetPrintAreasV1.js'
import type {NativeSheetViewportV2} from './nativeSheetGeometryV2.js'

export type NativeSheetPrintAreaSetV1 = {
 sheet_id:string;sheet_part:string;warnings:string[];
} & ({status:'available';areas:NativeSheetViewportV2[]}|{status:'unavailable';areas?:never})

/** Additive, ordered source ranges. Overlaps are refused, never merged or sorted. */
export function decodeNativeSheetPrintAreaSetsV1(input:unknown):NativeSheetPrintAreaSetV1[]{
 const value=snapshotNativePlainData(input,{maxDepth:7,maxNodes:16384})
 const fail=():never=>{throw new TypeError('Invalid native worksheet print area sets')}
 if(!Array.isArray(value)||value.length>64)return fail()
 const identities:unknown[]=[]
 const result=value.map(v=>{
  if(!v||typeof v!=='object'||Array.isArray(v))return fail()
  const o=v as Record<string,unknown>,available=o.status==='available'
  const keys=['sheet_id','sheet_part','status','warnings',...(available?['areas']:[])]
  if(Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k)))return fail()
  const base={sheet_id:o.sheet_id,sheet_part:o.sheet_part,status:o.status,warnings:o.warnings}
  if(!available){identities.push(base);return decodeNativeSheetPrintAreasV1([base])[0]! as NativeSheetPrintAreaSetV1}
  if(!Array.isArray(o.areas)||o.areas.length<1||o.areas.length>16)return fail()
  const areas=o.areas.map(area=>decodeNativeSheetPrintAreasV1([{...base,area}])[0]!.area!)
  identities.push({...base,area:areas[0]})
  for(let i=0;i<areas.length;i++)for(let j=0;j<i;j++){
   const a=areas[i]!,b=areas[j]!
   if(a.row<=b.end_row&&b.row<=a.end_row&&a.column<=b.end_column&&b.column<=a.end_column)return fail()
  }
  return {sheet_id:base.sheet_id as string,sheet_part:base.sheet_part as string,status:'available' as const,warnings:base.warnings as string[],areas}
 })
 decodeNativeSheetPrintAreasV1(identities)
 return result
}
