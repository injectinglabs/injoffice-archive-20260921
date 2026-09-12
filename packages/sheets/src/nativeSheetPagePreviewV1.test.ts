import {createRequire} from 'node:module'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {projectNativeWorkbookV2,createNativeMaximumDigitWidthAuthorityV2,compileNativeSheetGeometryV2,compileNativeStoredRowSheetGeometryV1,isCompiledNativeSheetGeometryV2,validateNativeSheetGeometryV2,compileNativeSheetPagePreviewV1,type NativeWorkbookV2,type NativeWorkbookObjectsV1} from './index.js'
const require=createRequire(import.meta.url)
function fixture(){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 const model=projectNativeWorkbookV2(workbook),font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
 const geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:2,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font))
 const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:geometry.source_package_sha256,tables:[],charts:[],page_settings:[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',warnings:['Source settings'],settings:{paper:'Letter',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}}]}
 return {geometry,objects,model,font}
}
describe('source-bound selected worksheet page geometry',()=>{
 it('rejects accessor and symbol host policies without invoking getters',()=>{
  const {geometry,objects}=fixture();let invoked=false
  const host={kind:'explicit-host-page-policy-v1',paper:'A4',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}
  Object.defineProperty(host,'scale',{enumerable:true,get(){invoked=true;return 100}})
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,host as any)).toThrow();expect(invoked).toBe(false)
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,{...objects,[Symbol('foreign')]:1} as any)).toThrow()
 })
 it('paginates whole bands down then over without dropping rows or columns',()=>{
  const {geometry,objects}=fixture()
  const p=compileNativeSheetPagePreviewV1(geometry,objects,{kind:'explicit-host-page-policy-v1',paper:'Letter',orientation:'portrait',scale:100,left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3})
  expect(p.pages).toHaveLength(9)
  expect(p.pages.map(p=>[p.rows.start,p.columns.start])).toEqual([[0,0],[1,0],[2,0],[0,1],[1,1],[2,1],[0,2],[1,2],[2,2]])
  for(const page of p.pages){expect(page.source_clip.x_emu*page.scale+page.translate_x_emu).toBe(page.content_clip.x_emu);expect(page.source_clip.y_emu*page.scale+page.translate_y_emu).toBe(page.content_clip.y_emu)}
 })
 it('brands stored-row approximation separately, preserves stored sizes and refuses stale evidence',()=>{
  const {model,font,objects}=fixture(),metric=createNativeMaximumDigitWidthAuthorityV2(model,font),viewport={row:0,column:0,end_row:2,end_column:2}
  objects.row_geometry=[{sheet_part:'Worksheets/Sheet1.xml',rows:Array.from({length:32},(_,row)=>({row,height_points:14.4,hidden:row===1})),warnings:['Stored only']}]
  const g=compileNativeStoredRowSheetGeometryV1(model,'7',viewport,metric,objects)
  expect(g.rows[0]!.height_emu).toBe(182880);expect(g.rows[1]!.height_emu).toBe(0)
  expect(g.approximation.policy).toBe('source-stored-rows-v1');expect(isCompiledNativeSheetGeometryV2(g)).toBe(false)
  expect(()=>validateNativeSheetGeometryV2(g)).toThrow()
  expect(compileNativeSheetPagePreviewV1(g,objects).warnings.join(' ')).toContain('baselines are not qualified')
  expect(()=>compileNativeStoredRowSheetGeometryV1(model,'7',{...viewport,end_row:32},metric,objects)).toThrow('first 32')
  objects.row_geometry[0]!.sheet_part='wrong.xml'
  expect(()=>compileNativeStoredRowSheetGeometryV1(model,'7',viewport,metric,objects)).toThrow('unavailable')
 })
 it('places stored geometry inside authored margins without changing source',()=>{
  const {geometry,objects}=fixture(),before=JSON.stringify(objects),p=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(p.settings_origin).toBe('source');expect(p.pages).toHaveLength(1)
  expect(p.pages[0]).toMatchObject({width_emu:7772400,height_emu:10058400,scale:1,translate_x_emu:914400,translate_y_emu:914400,source_clip:geometry.bounds})
  expect(JSON.stringify(objects)).toBe(before)
 })
 it('requires explicit host choice for unavailable settings and labels it',()=>{
  const {geometry,objects}=fixture();objects.page_settings![0]={sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'unavailable',warnings:['Printer settings not modeled']}
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects)).toThrow('explicit host')
  const p=compileNativeSheetPagePreviewV1(geometry,objects,{kind:'explicit-host-page-policy-v1',paper:'A4',orientation:'landscape',scale:50,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1})
  expect(p.settings_origin).toBe('explicit-host');expect(p.pages[0]).toMatchObject({width_emu:10692000,height_emu:7560000,scale:.5})
  expect(p.warnings.join(' ')).toContain('not authored')
 })
 it('refuses stale, forged geometry and an oversized column rather than clipping',()=>{
  const {geometry,objects}=fixture()
  expect(()=>compileNativeSheetPagePreviewV1({...geometry},objects)).toThrow('compiled')
  expect(()=>compileNativeSheetPagePreviewV1(geometry,{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
  objects.page_settings![0]!.settings!.left_inches=4;objects.page_settings![0]!.settings!.right_inches=4
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects)).toThrow('exceeds one page')
 })
})
