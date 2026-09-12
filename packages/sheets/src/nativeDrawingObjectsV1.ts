import {snapshotNativePlainData} from './nativePlainData.js'
export interface NativeDrawingMarkerV1 {column:number;row:number;column_offset_emu:number;row_offset_emu:number}
export type NativeDrawingAnchorV1={kind:'twoCellAnchor';from:NativeDrawingMarkerV1;to:NativeDrawingMarkerV1}|{kind:'oneCellAnchor';from:NativeDrawingMarkerV1;width_emu:number;height_emu:number}
export interface NativeDrawingObjectV1 {sheet_id:string;sheet_part:string;drawing_part:string;ordinal:number;kind:'chart'|'unsupported';chart_part?:string;anchor?:NativeDrawingAnchorV1;warnings:string[]}
export function decodeNativeDrawingObjectsV1(input:unknown):NativeDrawingObjectV1[]{
 const value=snapshotNativePlainData(input,{maxDepth:8,maxNodes:15000})
 const fail=():never=>{throw new TypeError('Invalid native source drawing objects')}
 const exact=(v:unknown,keys:string[])=>{if(!v||typeof v!=='object'||Array.isArray(v))return fail();const o=v as Record<string,unknown>;if(Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k)))return fail();return o}
 const int=(v:unknown,max:number,min=0):number=>typeof v==='number'&&Number.isSafeInteger(v)&&v>=min&&v<=max?v:fail()
 const part=(v:unknown,empty=false):string=>typeof v==='string'&&v.length<=1024&&((empty&&v==='')||(/^[A-Za-z0-9_.\/-]+$/.test(v)&&v.split('/').every(p=>p&&p!=='.'&&p!=='..')))?v:fail()
 const marker=(v:unknown):NativeDrawingMarkerV1=>{const m=exact(v,['column','row','column_offset_emu','row_offset_emu']);return {column:int(m.column,16383),row:int(m.row,1048575),column_offset_emu:int(m.column_offset_emu,2147483647),row_offset_emu:int(m.row_offset_emu,2147483647)}}
 if(!Array.isArray(value)||value.length>256)return fail()
 const seen=new Set<string>()
 return value.map(v=>{
  const base=v as Record<string,unknown>
  const o=exact(v,['sheet_id','sheet_part','drawing_part','ordinal','kind','warnings',...(base?.anchor!==undefined?['anchor']:[]),...(base?.chart_part!==undefined?['chart_part']:[])])
  if(typeof o.sheet_id!=='string'||!/^[1-9][0-9]{0,9}$/.test(o.sheet_id)||Number(o.sheet_id)>0xffffffff||(o.kind!=='chart'&&o.kind!=='unsupported')||!Array.isArray(o.warnings)||o.warnings.length<1||o.warnings.length>8||o.warnings.some(w=>typeof w!=='string'||w.length>4096))return fail()
  const sheet_part=part(o.sheet_part),drawing_part=part(o.drawing_part,true),ordinal=int(o.ordinal,256)
  const key=JSON.stringify([o.sheet_id,sheet_part,drawing_part,ordinal]);if(seen.has(key))return fail();seen.add(key)
  let anchor:NativeDrawingAnchorV1|undefined
  if(o.anchor!==undefined){
   const a=o.anchor as Record<string,unknown>
   if(a.kind==='twoCellAnchor'){exact(a,['kind','from','to']);anchor={kind:a.kind,from:marker(a.from),to:marker(a.to)}}
   else if(a.kind==='oneCellAnchor'){exact(a,['kind','from','width_emu','height_emu']);anchor={kind:a.kind,from:marker(a.from),width_emu:int(a.width_emu,2147483647,1),height_emu:int(a.height_emu,2147483647,1)}}else return fail()
  }
  if(o.kind==='chart'&&(!anchor||!drawing_part||ordinal<1||o.chart_part===undefined))return fail()
  if(o.kind==='unsupported'&&o.chart_part!==undefined)return fail()
  return {sheet_id:o.sheet_id,sheet_part,drawing_part,ordinal,kind:o.kind,warnings:o.warnings as string[],...(anchor?{anchor}:{}),...(o.chart_part===undefined?{}:{chart_part:part(o.chart_part)})}
 })
}
