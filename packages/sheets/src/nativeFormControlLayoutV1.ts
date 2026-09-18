import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {isCompiledNativeSheetGeometryV2,isCompiledNativeStoredRowSheetGeometryV1,type NativeSheetGeometryV2,type NativeSheetGeometryRectV2} from './nativeSheetGeometryV2.js'
import type {NativeFormControlV1} from './nativeFormControlObjectsV1.js'
import type {NativeDrawingMarkerV1} from './nativeDrawingObjectsV1.js'

const EMU_PER_POINT=12700

/**
 * Excel paints a checkbox form control as a fixed 12 pt square inside the
 * control's own anchor rectangle, at a fixed offset from that rectangle's
 * top-left corner. The three constants below are read from Excel 16.112.4's
 * PDF export of the hard-v2 corpus file checkbox-form-control-align.xlsx,
 * whose two controls sit at different rows and carry different caption
 * alignments:
 *
 *   65.5 718.5 12 12 re f   (white fill)  and  37.5 28.5 12 12 re S  (black
 *   stroke under the flip 1 0 0 -1 28 759 cm), then the same pair 18 pt lower.
 *
 * Both squares are exactly 12 x 12 pt, both sit 5.35 pt right of and 3 pt below
 * their control's anchored top-left corner, and neither carries a check mark:
 * their ctrlProps state no `checked`, so the box is empty. The offset is the
 * same for both although their captions are aligned differently, so the square
 * follows the control rectangle and not the caption.
 */
const CHECKBOX_SIDE_POINTS=12
const CHECKBOX_LEFT_INSET_POINTS=5.35
const CHECKBOX_TOP_INSET_POINTS=3

/**
 * Excel's legacy-control text insets, the values it writes into the DrawingML
 * mirror of a control whose VML states o:insetmode="auto": 27432 EMU (2.16 pt)
 * left and right, 22860 EMU (1.8 pt) top and bottom.
 */
const CAPTION_INSET_HORIZONTAL_EMU=27432
const CAPTION_INSET_VERTICAL_EMU=22860

export interface NativeFormControlCaptionBoxV1 {
 readonly text:string
 readonly size_points:number
 readonly align:'left'|'center'|'right'
 readonly valign:'top'|'center'|'bottom'
 readonly rect:NativeSheetGeometryRectV2
}
export interface NativePositionedFormControlV1 {
 readonly source:NativeFormControlV1
 readonly status:'positioned'|'outside'|'unavailable'
 /** The control's whole anchored rectangle, in viewport-local EMU. */
 readonly rect?:NativeSheetGeometryRectV2
 /** The 12 pt box to stroke and fill. Present only when status is 'positioned'. */
 readonly box?:NativeSheetGeometryRectV2
 /** Whether the box carries a check mark, from the control's saved state. */
 readonly checked?:boolean
 readonly caption?:NativeFormControlCaptionBoxV1
 readonly clip?:NativeSheetGeometryRectV2
 readonly warning?:string
}

/**
 * Source-anchored placement for the worksheet's checkbox form controls, in
 * viewport-local EMU. Only a control the source states completely as a
 * checkbox is positioned; every other ObjectType is reported unavailable with
 * the reason, and no box or caption is placed for it. The caption and the box
 * are placed together or not at all: a caption with no box would be a control
 * that is not there.
 */
export function layoutNativeFormControlsV1(geometry:NativeSheetGeometryV2,objects:NativeWorkbookObjectsV1):NativePositionedFormControlV1[]{
 if(!isCompiledNativeSheetGeometryV2(geometry)&&!isCompiledNativeStoredRowSheetGeometryV1(geometry))throw new TypeError('Form control layout requires compiled source geometry')
 const source=decodeNativeWorkbookObjectsV1(objects,geometry.source_package_sha256)
 const page=source.page_settings?.find(s=>s.sheet_id===geometry.sheet_id)
 if(!page)throw new TypeError('Form control source worksheet identity unavailable')
 const rows=new Map(geometry.rows.map(r=>[r.row,r])),cols=new Map(geometry.columns.map(c=>[c.column,c]))
 const point=(m:NativeDrawingMarkerV1)=>{
  const row=rows.get(m.row),col=cols.get(m.column)
  if(!row||!col||m.row_offset_emu>row.height_emu||m.column_offset_emu>col.width_emu)return undefined
  return {x:col.x_emu+m.column_offset_emu,y:row.y_emu+m.row_offset_emu}
 }
 return (source.form_controls??[]).filter(c=>c.sheet_id===geometry.sheet_id).map(control=>{
  const unavailable=(warning:string):NativePositionedFormControlV1=>({source:control,status:'unavailable',warning})
  if(control.sheet_part!==page.sheet_part)return unavailable('Form control worksheet part does not join the selected geometry')
  if(control.kind!=='checkbox')return unavailable(control.warnings[0]??'Form control is not a checkbox this tier paints; no box is drawn')
  const anchor=control.anchor
  if(!anchor)return unavailable('Source anchor is unqualified; no position is guessed')
  const start=point(anchor.from)
  if(!start)return unavailable('Anchor marker lies outside calibrated source geometry or its stored cell extent')
  const end=point(anchor.to)
  if(!end)return unavailable('Anchor endpoint lies outside calibrated source geometry or its stored cell extent')
  if(end.x<=start.x||end.y<=start.y)return unavailable('Source anchor does not define a positive rectangle')
  const rect={x_emu:start.x,y_emu:start.y,width_emu:end.x-start.x,height_emu:end.y-start.y}
  const side=CHECKBOX_SIDE_POINTS*EMU_PER_POINT
  const box={
   x_emu:rect.x_emu+CHECKBOX_LEFT_INSET_POINTS*EMU_PER_POINT,
   y_emu:rect.y_emu+CHECKBOX_TOP_INSET_POINTS*EMU_PER_POINT,
   width_emu:side,height_emu:side,
  }
  // A control smaller than Excel's own glyph would paint a box outside the
  // rectangle the source anchored. That is not a checkbox at a known position,
  // so it refuses rather than overflowing.
  if(box.x_emu+box.width_emu>rect.x_emu+rect.width_emu||box.y_emu+box.height_emu>rect.y_emu+rect.height_emu){
   return unavailable('Source anchor rectangle is smaller than the control glyph Excel paints; no box is placed outside it')
  }
  const captionWidth=rect.width_emu-2*CAPTION_INSET_HORIZONTAL_EMU
  const captionHeight=rect.height_emu-2*CAPTION_INSET_VERTICAL_EMU
  if(captionWidth<=0||captionHeight<=0)return unavailable('Source anchor rectangle leaves no caption area inside the control text insets; no box is placed without its caption')
  const caption:NativeFormControlCaptionBoxV1={
   text:control.caption!,size_points:control.caption_size_points!,
   align:control.caption_align!,valign:control.caption_valign!,
   rect:{x_emu:rect.x_emu+CAPTION_INSET_HORIZONTAL_EMU,y_emu:rect.y_emu+CAPTION_INSET_VERTICAL_EMU,width_emu:captionWidth,height_emu:captionHeight},
  }
  const x=Math.max(0,start.x),y=Math.max(0,start.y)
  const right=Math.min(geometry.bounds.width_emu,end.x),bottom=Math.min(geometry.bounds.height_emu,end.y)
  if(right<=x||bottom<=y)return {source:control,status:'outside'}
  return {source:control,status:'positioned',rect,box,checked:control.checked,caption,clip:{x_emu:x,y_emu:y,width_emu:right-x,height_emu:bottom-y}}
 })
}
