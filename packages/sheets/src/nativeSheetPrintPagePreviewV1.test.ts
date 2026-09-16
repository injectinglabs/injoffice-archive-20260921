import {createRequire} from 'node:module'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {
  compileNativeSheetGeometryV2,
  compileNativeSheetPrintPagePreviewV1,
  createNativeMaximumDigitWidthAuthorityV2,
  EMU_PER_CSS_PIXEL,
  NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI,
  NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_PROTOCOL,
  projectNativeWorkbookV2,
  type NativeSheetViewportV2,
  type NativeWorkbookObjectsV1,
  type NativeWorkbookV2,
} from './index.js'
const require=createRequire(import.meta.url)
const area:NativeSheetViewportV2={row:1,column:1,end_row:3,end_column:3}
function fixture(){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 const model=projectNativeWorkbookV2(workbook),font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
 const part=model.sheets[0]!.mutation_authority.source_part
 const objects:NativeWorkbookObjectsV1={
  protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:model.source.package_sha256,tables:[],charts:[],
  print_area_sets:[{sheet_id:'7',sheet_part:part,status:'available',areas:[{...area}],warnings:['Saved source rectangle']}],
  page_settings:[{sheet_id:'7',sheet_part:part,status:'available',warnings:['Source settings'],settings:{paper:'Letter',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}}],
 }
 const geometry=compileNativeSheetGeometryV2(model,'7',area,createNativeMaximumDigitWidthAuthorityV2(model,font))
 return {model,objects,geometry,part}
}
describe('source-only 96 DPI print-page preview',()=>{
 it('returns one saved-area Letter page with exact 96 DPI capture rectangles',()=>{
  const {objects,geometry}=fixture(),before=JSON.stringify(objects)
  const preview=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(preview.status).toBe('available')
  if(preview.status!=='available')return
  expect(preview.protocol).toBe(NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_PROTOCOL)
  expect(preview.dpi).toBe(NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI)
  expect(preview.dpi).toBe(96)
  expect(preview.read_only).toBe(true)
  expect(preview.fidelity).toBe('approximate')
  expect(preview.settings_origin).toBe('source')
  expect(preview.pages).toHaveLength(1)
  const page=preview.pages[0]!
  expect(page.sequence).toBe(1)
  expect(page.number).toBe(1)
  expect(page.area_index).toBe(0)
  expect(page.width_emu).toBe(7_772_400)
  expect(page.height_emu).toBe(10_058_400)
  expect(page.width_css_px).toBe(816)
  expect(page.height_css_px).toBe(1056)
  expect(page.width_css_px).toBe(page.width_emu/EMU_PER_CSS_PIXEL)
  expect(page.content_clip_css_px).toEqual({x_css_px:96,y_css_px:96,width_css_px:624,height_css_px:864})
  expect(page.rows).toEqual({start:1,end:3})
  expect(page.columns).toEqual({start:1,end:3})
  expect(page.scale).toBe(1)
  expect(page.paint).toBeUndefined()
  expect(page.source_clip_css_px.x_css_px*page.scale+page.translate_x_css_px).toBe(page.content_clip_css_px.x_css_px)
  expect(page.source_clip_css_px.y_css_px*page.scale+page.translate_y_css_px).toBe(page.content_clip_css_px.y_css_px)
  expect(preview.warnings.join(' ')).toContain('not Excel printer calibration')
  expect(preview.warnings.join(' ')).toContain('fit-to-page qualification')
  expect(JSON.stringify(objects)).toBe(before)
  expect(Object.isFrozen(preview)).toBe(true)
  expect(Object.isFrozen(preview.pages)).toBe(true)
 })
 it('refuses printer-dependent and missing page setup without inventing paper or dropping identity',()=>{
  const {objects,geometry,part}=fixture()
  objects.page_settings=[{sheet_id:'7',sheet_part:part,status:'unavailable',warnings:['Worksheet page preview requires explicit bounded paper, orientation, margins and percentage scale or fit-to-page settings; printer defaults and unsupported print options are not inferred.']}]
  const printer=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(printer.status).toBe('unavailable')
  if(printer.status!=='unavailable')return
  expect(printer.pages).toEqual([])
  expect(printer.reason).toContain('printer defaults')
  expect(printer.reason).toContain('does not invent paper, margins, scale')
  expect(printer.reason).toContain('Grid preview is unchanged.')
  expect(printer.settings_origin).toBe('source')
  expect(printer.document_id).toBe(geometry.document_id)
  expect(printer.sheet_id).toBe(geometry.sheet_id)
  delete objects.page_settings
  const missing=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(missing.status).toBe('unavailable')
  if(missing.status!=='unavailable')return
  expect(missing.reason).toContain('Source page settings are missing')
  expect(missing.reason).toContain('Grid preview is unchanged.')
  expect(missing.pages).toHaveLength(0)
 })
 it('refuses fit-to-page, missing saved areas, and host-style option extras',()=>{
  const {objects,geometry,part}=fixture()
  Object.assign(objects.page_settings![0]!.settings!,{fit_to_page:{width:1,height:1}})
  const fit=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(fit.status).toBe('unavailable')
  if(fit.status!=='unavailable')return
  expect(fit.reason).toContain('fit-to-page')
  expect(fit.reason).toContain('does not invent a fit scale')
  expect(fit.pages).toEqual([])
  delete objects.page_settings![0]!.settings!.fit_to_page
  objects.print_area_sets=[{sheet_id:'7',sheet_part:part,status:'unavailable',warnings:['Unsupported print area']}]
  const area=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(area.status).toBe('unavailable')
  if(area.status!=='unavailable')return
  expect(area.reason).toContain('Unsupported print area')
  expect(area.reason).toContain('used-range or A1 fallback')
  delete objects.print_area_sets
  const missingArea=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(missingArea.status).toBe('unavailable')
  if(missingArea.status!=='unavailable')return
  expect(missingArea.reason).toContain('Saved print area does not join')
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects,{repeat_print_titles:true,extra:1} as never)).toThrow('explicit plain data')
  let called=false
  const accessor={repeat_print_titles:true as const}
  Object.defineProperty(accessor,'repeat_print_titles',{enumerable:true,get(){called=true;return true}})
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects,accessor)).toThrow();expect(called).toBe(false)
 })
 it('rejects uncompiled geometry, stale objects, and mismatched paint plans rather than guessing',()=>{
  const {objects,geometry}=fixture()
  expect(()=>compileNativeSheetPrintPagePreviewV1([{...geometry}],objects)).toThrow('compiled source-qualified')
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects,{paint_plans:[]})).toThrow('paint plans')
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects,{paint_plans:[{geometry_sha256:`sha256:${'c'.repeat(64)}`,sheet_id:'7',source_package_sha256:geometry.source_package_sha256}]} as never)).toThrow('join compiled geometry')
 })
})

describe('host-default paper for a worksheet that authors margins and no pageSetup',()=>{
 const marginsOnly=(part:string)=>({
  sheet_id:'7',sheet_part:part,status:'margins-only' as const,
  warnings:['Worksheet declares authored page margins and no pageSetup element, so no paper size, orientation or scale is authored. Page geometry requires a host paper choice; this tier does not select one.'],
  margins:{left_inches:1,right_inches:1,top_inches:1,bottom_inches:1},
 })
 it('paginates at the host default and says so, matching the authored-Letter geometry exactly',()=>{
  const {objects,geometry,part}=fixture()
  const authored=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  objects.page_settings=[marginsOnly(part)]
  const preview=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(preview.status).toBe('available')
  if(preview.status!=='available'||authored.status!=='available')return
  expect(preview.settings_origin).toBe('host-default')
  expect(preview.warnings.some(w=>w.includes('host default, not authored workbook settings'))).toBe(true)
  // US Letter portrait at 96 DPI, identical to the same workbook with the paper authored.
  expect(preview.pages).toHaveLength(authored.pages.length)
  expect(preview.pages[0]!.width_css_px).toBe(816)
  expect(preview.pages[0]!.height_css_px).toBe(1056)
  expect(preview.pages.map(p=>[p.width_emu,p.height_emu])).toEqual(authored.pages.map(p=>[p.width_emu,p.height_emu]))
 })
 it('keeps refusing an authored pageSetup this tier cannot support, and margins-only with no margins',()=>{
  const {objects,geometry,part}=fixture()
  objects.page_settings=[{sheet_id:'7',sheet_part:part,status:'unavailable',warnings:['Unsupported page settings']}]
  const unsupported=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(unsupported.status).toBe('unavailable')
  if(unsupported.status!=='unavailable')return
  expect(unsupported.reason).toContain('Unsupported page settings')
  expect(unsupported.settings_origin).toBe('source')
  // A margins-only record with no margins is a malformed payload, not a soft
  // refusal: the decoder requires the key for that status and rejects it.
  const {margins:_omitted,...withoutMargins}=marginsOnly(part)
  objects.page_settings=[withoutMargins as never]
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects)).toThrow('Invalid native worksheet page settings')
 })
})
