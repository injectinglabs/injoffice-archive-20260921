import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {isCompiledNativeSheetGeometryV2,isCompiledNativeStoredRowSheetGeometryV1,type NativeSheetGeometryV2,type NativeSheetGeometryRectV2} from './nativeSheetGeometryV2.js'
import type {NativeDrawingObjectV1,NativeDrawingMarkerV1} from './nativeDrawingObjectsV1.js'
export interface NativePositionedDrawingV1 {
 source:NativeDrawingObjectV1;status:'positioned'|'outside'|'unavailable';rect?:NativeSheetGeometryRectV2;clip?:NativeSheetGeometryRectV2;warning?:string;
}
/** Coordinates are viewport-local EMU. Unknown/off-viewport metrics are never
 * inferred, and the chart's own cached plot remains a separate approximation.
 */
export function layoutNativeDrawingObjectsV1(geometry:NativeSheetGeometryV2,objects:NativeWorkbookObjectsV1):NativePositionedDrawingV1[]{
 if(!isCompiledNativeSheetGeometryV2(geometry)&&!isCompiledNativeStoredRowSheetGeometryV1(geometry))throw new TypeError('Drawing layout requires compiled source geometry')
 const source=decodeNativeWorkbookObjectsV1(objects,geometry.source_package_sha256)
 const page=source.page_settings?.find(s=>s.sheet_id===geometry.sheet_id)
 if(!page)throw new TypeError('Drawing source worksheet identity unavailable')
 const rows=new Map(geometry.rows.map(r=>[r.row,r])),cols=new Map(geometry.columns.map(c=>[c.column,c]))
 const point=(m:NativeDrawingMarkerV1)=>{
  const row=rows.get(m.row),col=cols.get(m.column)
  if(!row||!col||m.row_offset_emu>row.height_emu||m.column_offset_emu>col.width_emu)return undefined
  return {x:col.x_emu+m.column_offset_emu,y:row.y_emu+m.row_offset_emu}
 }
 return (source.drawing_objects??[]).filter(d=>d.sheet_id===geometry.sheet_id).map(d=>{
  const unavailable=(warning:string):NativePositionedDrawingV1=>({source:d,status:'unavailable',warning})
  if(d.sheet_part!==page.sheet_part)return unavailable('Drawing worksheet part does not join the selected geometry')
  if(!d.anchor)return unavailable('Source anchor is unqualified; no position is guessed')
  if(d.kind==='chart'&&!source.charts.some(c=>c.part===d.chart_part))return unavailable('Source chart cache part is not present in this inspection')
  const a=d.anchor,start=point(a.from)
  if(!start)return unavailable('Anchor marker lies outside calibrated source geometry or its stored cell extent')
  const end=a.kind==='twoCellAnchor'?point(a.to):{x:start.x+a.width_emu,y:start.y+a.height_emu}
  if(!end)return unavailable('Anchor endpoint lies outside calibrated source geometry or its stored cell extent')
  if(end.x<=start.x||end.y<=start.y)return unavailable('Source anchor does not define a positive rectangle')
  const rect={x_emu:start.x,y_emu:start.y,width_emu:end.x-start.x,height_emu:end.y-start.y}
  const x=Math.max(0,start.x),y=Math.max(0,start.y),right=Math.min(geometry.bounds.width_emu,end.x),bottom=Math.min(geometry.bounds.height_emu,end.y)
  if(right<=x||bottom<=y)return {source:d,status:'outside'}
  return {source:d,status:'positioned',rect,clip:{x_emu:x,y_emu:y,width_emu:right-x,height_emu:bottom-y}}
 })
}
