import {createRequire} from 'node:module'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {projectNativeWorkbookV2,createNativeMaximumDigitWidthAuthorityV2,compileNativeSheetGeometryV2,layoutNativeFormControlsV1,type NativeWorkbookV2,type NativeWorkbookObjectsV1,type NativeFormControlV1} from './index.js'
import {decodeNativeFormControlsV1} from './nativeFormControlObjectsV1.js'
const require=createRequire(import.meta.url)

const EMU_PER_POINT=12700
/** The two markers of Check Box 1 in the hard-v2 corpus file
 * checkbox-form-control-align.xlsx, verbatim. */
const CHECKBOX:NativeFormControlV1={
 sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',ordinal:1,shape_id:'1025',name:'Check Box 1',
 kind:'checkbox',object_type:'CheckBox',checked:false,
 anchor:{kind:'twoCellAnchor',from:{column:0,row:0,column_offset_emu:123825,row_offset_emu:57150},to:{column:2,row:1,column_offset_emu:428625,row_offset_emu:85725}},
 caption:'All effects',caption_size_points:8,caption_align:'right',caption_valign:'bottom',
 control_part:'xl/ctrlProps/ctrlProp1.xml',legacy_part:'xl/drawings/vmlDrawing1.vml',
 warnings:['Source anchor, control type and legacy caption only.'],
}
function fixture(controls:NativeFormControlV1[]=[structuredClone(CHECKBOX)]){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 const model=projectNativeWorkbookV2(workbook),font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
 const geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:2,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font))
 const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:geometry.source_package_sha256,tables:[],charts:[],page_settings:[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'unavailable',warnings:['No print settings']}],form_controls:controls}
 return {geometry,objects}
}

describe('source-positioned form control layout',()=>{
 it('places the checkbox glyph where Excel places it',()=>{
  // Excel 16.112.4's own PDF export of checkbox-form-control-align.xlsx paints
  // `65.5 718.5 12 12 re f` for this control on US Letter with 0.7 in left and
  // 0.75 in top margins, so the box is 12 x 12 pt with its top-left corner
  // 65.5 pt from the page's left edge and 792 - 718.5 - 12 = 61.5 pt from its
  // top. Subtracting the margins (50.4 pt and 54 pt) leaves 15.1 pt and 7.5 pt
  // inside the sheet, which is what this layout must produce.
  const {geometry,objects}=fixture(),before=JSON.stringify(objects)
  const [placed]=layoutNativeFormControlsV1(geometry,objects)
  expect(placed!.status).toBe('positioned')
  expect(placed!.box).toEqual({x_emu:15.1*EMU_PER_POINT,y_emu:7.5*EMU_PER_POINT,width_emu:12*EMU_PER_POINT,height_emu:12*EMU_PER_POINT})
  expect(placed!.rect).toEqual({x_emu:123825,y_emu:57150,width_emu:geometry.columns[2]!.x_emu+428625-123825,height_emu:geometry.rows[1]!.y_emu+85725-57150})
  // Excel's ctrlProps state no `checked`, and its PDF draws a stroked empty
  // square with no mark. Nothing here invents one.
  expect(placed!.checked).toBe(false)
  expect(JSON.stringify(objects)).toBe(before)
 })
 it('carries the caption into the control text area with its authored alignment',()=>{
  const {geometry,objects}=fixture()
  const caption=layoutNativeFormControlsV1(geometry,objects)[0]!.caption!
  expect(caption.text).toBe('All effects')
  expect(caption.size_points).toBe(8)
  expect(caption.align).toBe('right')
  expect(caption.valign).toBe('bottom')
  // Excel's legacy-control insets: 27432 EMU left and right, 22860 top and bottom.
  expect(caption.rect).toEqual({x_emu:123825+27432,y_emu:57150+22860,width_emu:geometry.columns[2]!.x_emu+428625-123825-2*27432,height_emu:geometry.rows[1]!.y_emu+85725-57150-2*22860})
 })
 it('paints nothing for a control that is not a checkbox',()=>{
  for(const objectType of ['Button','Drop','Radio','Spin']){
   const source={...structuredClone(CHECKBOX),kind:'unsupported' as const,object_type:objectType,warnings:[`Form control ${JSON.stringify(objectType)} is not a checkbox.`]}
   const {geometry,objects}=fixture([source])
   const [placed]=layoutNativeFormControlsV1(geometry,objects)
   expect(placed!.status).toBe('unavailable')
   expect(placed!.box).toBeUndefined()
   expect(placed!.caption).toBeUndefined()
   expect(placed!.warning).toContain(objectType)
  }
 })
 it('refuses a control rectangle too small to hold the glyph rather than overflowing it',()=>{
  const source=structuredClone(CHECKBOX)
  source.anchor={kind:'twoCellAnchor',from:{column:0,row:0,column_offset_emu:0,row_offset_emu:0},to:{column:0,row:0,column_offset_emu:60000,row_offset_emu:60000}}
  const {geometry,objects}=fixture([source])
  const [placed]=layoutNativeFormControlsV1(geometry,objects)
  expect(placed!.status).toBe('unavailable')
  expect(placed!.box).toBeUndefined()
  expect(placed!.warning).toContain('smaller than the control glyph')
 })
 it('refuses forged geometry, stale packages and wrong worksheet parts',()=>{
  const {geometry,objects}=fixture()
  expect(()=>layoutNativeFormControlsV1({...geometry},objects)).toThrow('compiled')
  expect(()=>layoutNativeFormControlsV1(geometry,{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
  objects.form_controls![0]!.sheet_part='wrong.xml'
  expect(layoutNativeFormControlsV1(geometry,objects)[0]!.status).toBe('unavailable')
 })
 it('never decodes a checkbox missing a fact the painter reads',()=>{
  expect(decodeNativeFormControlsV1([structuredClone(CHECKBOX)])).toHaveLength(1)
  for(const key of ['anchor','caption','caption_size_points','caption_align','caption_valign'] as const){
   const broken=structuredClone(CHECKBOX) as unknown as Record<string,unknown>
   delete broken[key]
   expect(()=>decodeNativeFormControlsV1([broken])).toThrow()
  }
  const duplicate=structuredClone(CHECKBOX)
  expect(()=>decodeNativeFormControlsV1([duplicate,structuredClone(CHECKBOX)])).toThrow()
  const accessor=structuredClone(CHECKBOX) as unknown as Record<string,unknown>
  let called=false;Object.defineProperty(accessor,'ordinal',{enumerable:true,get(){called=true;return 1}})
  expect(()=>decodeNativeFormControlsV1([accessor])).toThrow();expect(called).toBe(false)
 })
})
