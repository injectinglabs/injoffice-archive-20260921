import {createRequire} from 'node:module'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {projectNativeWorkbookV2,createNativeMaximumDigitWidthAuthorityV2,compileNativeSheetGeometryV2,layoutNativeDrawingObjectsV1,type NativeWorkbookV2,type NativeWorkbookObjectsV1} from './index.js'
import {decodeNativeDrawingObjectsV1} from './nativeDrawingObjectsV1.js'
const require=createRequire(import.meta.url)
function fixture(){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 const model=projectNativeWorkbookV2(workbook),font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
 const geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:2,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font))
 const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:geometry.source_package_sha256,tables:[],charts:[],page_settings:[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'unavailable',warnings:['No print settings']}],drawing_objects:[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',drawing_part:'Drawings/drawing.xml',ordinal:1,kind:'unsupported',warnings:['Unknown drawing retained'],anchor:{kind:'twoCellAnchor',from:{column:0,row:0,column_offset_emu:100,row_offset_emu:200},to:{column:2,row:2,column_offset_emu:300,row_offset_emu:400}}}]}
 return {geometry,objects}
}
describe('source-positioned drawing layout',()=>{
 it('retains unknown source objects and exact offsets without changing input',()=>{
  const {geometry,objects}=fixture(),before=JSON.stringify(objects),[p]=layoutNativeDrawingObjectsV1(geometry,objects)
  expect(p!.status).toBe('positioned');expect(p!.source.kind).toBe('unsupported')
  expect(p!.rect).toEqual({x_emu:100,y_emu:200,width_emu:geometry.columns[2]!.x_emu+200,height_emu:geometry.rows[2]!.y_emu+200})
  expect(p!.clip).toEqual(p!.rect);expect(JSON.stringify(objects)).toBe(before)
 })
 it('clips explicit one-cell extents but never guesses absent endpoint geometry',()=>{
  const {geometry,objects}=fixture(),d=objects.drawing_objects![0]!
  d.anchor={kind:'oneCellAnchor',from:{column:0,row:0,column_offset_emu:0,row_offset_emu:0},width_emu:2000000000,height_emu:2000000000}
  expect(layoutNativeDrawingObjectsV1(geometry,objects)[0]!.clip).toEqual(geometry.bounds)
  d.anchor={kind:'twoCellAnchor',from:d.anchor.from,to:{column:3,row:2,column_offset_emu:0,row_offset_emu:0}}
  expect(layoutNativeDrawingObjectsV1(geometry,objects)[0]!.status).toBe('unavailable')
 })
 it('refuses forged geometry, stale packages, wrong worksheet parts and missing chart ownership',()=>{
  const {geometry,objects}=fixture()
  expect(()=>layoutNativeDrawingObjectsV1({...geometry},objects)).toThrow('compiled')
  expect(()=>layoutNativeDrawingObjectsV1(geometry,{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
  objects.drawing_objects![0]!.sheet_part='wrong.xml'
  expect(layoutNativeDrawingObjectsV1(geometry,objects)[0]!.status).toBe('unavailable')
  Object.assign(objects.drawing_objects![0],{sheet_part:'Worksheets/Sheet1.xml',kind:'chart',chart_part:'Charts/missing.xml'})
  expect(layoutNativeDrawingObjectsV1(geometry,objects)[0]!.warning).toContain('cache part')
 })
 it('rejects duplicate identity, malformed anchors and accessor payloads',()=>{
  const {objects}=fixture(),d=objects.drawing_objects![0]!
  expect(()=>decodeNativeDrawingObjectsV1([d,d])).toThrow()
  expect(()=>decodeNativeDrawingObjectsV1([{...d,anchor:{...d.anchor,extra:1}}])).toThrow()
  let called=false;Object.defineProperty(d,'ordinal',{enumerable:true,get(){called=true;return 1}})
  expect(()=>decodeNativeDrawingObjectsV1([d])).toThrow();expect(called).toBe(false)
 })
})
