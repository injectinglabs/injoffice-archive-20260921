import {createRequire} from 'node:module'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {decodeNativeSheetPrintAreasV1,selectNativeSheetPrintAreaV1,decodeNativeWorkbookObjectsV1,projectNativeWorkbookV2,compileNativeSheetGeometryV2,createNativeMaximumDigitWidthAuthorityV2,layoutNativeDrawingObjectsV1,type NativeWorkbookV2,type NativeWorkbookObjectsV1} from './index.js'
const require=createRequire(import.meta.url)
function fixture(){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 const model=projectNativeWorkbookV2(workbook)
 const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:model.source.package_sha256,tables:[],charts:[],print_areas:[{sheet_id:'7',sheet_part:model.sheets[0]!.mutation_authority.source_part,status:'available',area:{row:1,column:1,end_row:3,end_column:3},warnings:['Bounded saved source rectangle']}]}
 return {model,objects}
}
describe('bounded saved print areas',()=>{
 it('owns a plain copy and retains absent optional metadata',()=>{
  const {objects}=fixture(),copy=decodeNativeWorkbookObjectsV1(objects,objects.package_sha256)
  expect(copy).toEqual(objects);copy.print_areas![0]!.warnings[0]='changed'
  expect(objects.print_areas![0]!.warnings[0]).not.toBe('changed')
  delete objects.print_areas
  expect(Object.hasOwn(decodeNativeWorkbookObjectsV1(objects,objects.package_sha256),'print_areas')).toBe(false)
  expect(()=>decodeNativeWorkbookObjectsV1({...objects,print_areas:undefined},objects.package_sha256)).toThrow()
 })
 it('requires package, worksheet ID/part and a branded model',()=>{
  const {model,objects}=fixture()
  expect(selectNativeSheetPrintAreaV1(model,'7',objects)).toEqual({row:1,column:1,end_row:3,end_column:3})
  expect(()=>selectNativeSheetPrintAreaV1({...model},'7',objects)).toThrow('projected')
  expect(()=>selectNativeSheetPrintAreaV1(model,'8',objects)).toThrow('join')
  expect(()=>selectNativeSheetPrintAreaV1(model,'7',{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
  objects.print_areas![0]!.sheet_part='wrong.xml'
  expect(()=>selectNativeSheetPrintAreaV1(model,'7',objects)).toThrow('join')
 })
 it('does not fall back for absent or unavailable saved areas',()=>{
  const {model,objects}=fixture(),entry=objects.print_areas![0]!
  objects.print_areas=[{sheet_id:entry.sheet_id,sheet_part:entry.sheet_part,status:'unavailable',warnings:['No supported area']}]
  expect(decodeNativeSheetPrintAreasV1(objects.print_areas)).toEqual(objects.print_areas)
  expect(()=>selectNativeSheetPrintAreaV1(model,'7',objects)).toThrow('unavailable')
  objects.print_areas=[];expect(()=>selectNativeSheetPrintAreaV1(model,'7',objects)).toThrow('join')
  delete objects.print_areas;expect(()=>selectNativeSheetPrintAreaV1(model,'7',objects)).toThrow('join')
 })
 it('rejects duplicate identity, extra keys, unavailable area, non-plain prototypes and accessors',()=>{
  const entry=fixture().objects.print_areas![0]!
  for(const entries of [[entry,entry],[entry,{...entry,sheet_id:'8'}],[entry,{...entry,sheet_part:'other.xml'}],[{...entry,extra:1}],[{...entry,status:'unavailable'}],[Object.assign(Object.create({inherited:true}),entry)],Array(65).fill(entry)])expect(()=>decodeNativeSheetPrintAreasV1(entries)).toThrow()
  let called=false;const getter={...entry};Object.defineProperty(getter,'area',{enumerable:true,get(){called=true;return entry.area}})
  expect(()=>decodeNativeSheetPrintAreasV1([getter])).toThrow();expect(called).toBe(false)
  for(const warnings of [[],Array(9).fill('warning'),['a'.repeat(4097)]])expect(()=>decodeNativeSheetPrintAreasV1([{...entry,warnings}])).toThrow()
 })
 it('rejects reversed/noninteger/Excel-exceeding rectangles and preserves geometry budgets without clipping',()=>{
  const {model,objects}=fixture(),entry=objects.print_areas![0]!
  for(const area of [{row:-0,column:0,end_row:1,end_column:1},{row:-1,column:0,end_row:1,end_column:1},{row:0.5,column:0,end_row:1,end_column:1},{row:2,column:0,end_row:1,end_column:1},{row:0,column:2,end_row:1,end_column:1},{row:0,column:0,end_row:1048576,end_column:1},{row:0,column:0,end_row:1,end_column:16384},{row:0,column:0,end_row:Infinity,end_column:1},{row:0,column:0,end_row:1,end_column:1,extra:1}])expect(()=>decodeNativeSheetPrintAreasV1([{...entry,area}])).toThrow()
  for(const area of [{row:0,column:0,end_row:4096,end_column:0},{row:0,column:0,end_row:0,end_column:1024},{row:0,column:0,end_row:999,end_column:100}]){
   objects.print_areas=[{...entry,status:'available',area}]
   expect(()=>decodeNativeSheetPrintAreasV1(objects.print_areas)).not.toThrow()
   expect(()=>selectNativeSheetPrintAreaV1(model,'7',objects)).toThrow('viewport limits')
  }
 })
 it('retains a non-A1 origin and source chart anchor offsets through geometry',()=>{
  const {model,objects}=fixture(),entry=objects.print_areas![0]!
  const font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
  const geometry=compileNativeSheetGeometryV2(model,'7',selectNativeSheetPrintAreaV1(model,'7',objects),createNativeMaximumDigitWidthAuthorityV2(model,font))
  expect(geometry.origin_cell).toEqual({row:1,column:1});expect(geometry.rows[0]!.y_emu).toBe(0);expect(geometry.columns[0]!.x_emu).toBe(0)
  objects.page_settings=[{sheet_id:'7',sheet_part:entry.sheet_part,status:'unavailable',warnings:['No settings']}]
  objects.charts=[{part:'Charts/chart1.xml',type:'col',series:[],warnings:[]}]
  objects.drawing_objects=[{sheet_id:'7',sheet_part:entry.sheet_part,drawing_part:'Drawings/drawing1.xml',ordinal:1,kind:'chart',chart_part:'Charts/chart1.xml',warnings:['Source anchor'],anchor:{kind:'twoCellAnchor',from:{row:1,column:1,row_offset_emu:200,column_offset_emu:100},to:{row:3,column:3,row_offset_emu:400,column_offset_emu:300}}}]
  const [drawing]=layoutNativeDrawingObjectsV1(geometry,objects)
  expect(drawing!.status).toBe('positioned');expect(drawing!.rect).toEqual({x_emu:100,y_emu:200,width_emu:geometry.columns[2]!.x_emu+200,height_emu:geometry.rows[2]!.y_emu+200})
 })
})
