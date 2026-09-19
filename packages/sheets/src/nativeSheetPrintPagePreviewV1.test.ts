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
 it('reserves the authored header and footer bands of a margins-only worksheet',()=>{
  // The host chooses paper for such a worksheet; the six authored margins are
  // still the worksheet's own, and the header band is measured from the paper
  // edge, so a header deeper than the top margin lowers the body.
  const {objects,geometry,part}=fixture()
  objects.page_settings=[{...marginsOnly(part),margins:{left_inches:1,right_inches:1,top_inches:1,bottom_inches:1,header_inches:1.5,footer_inches:0.3}}]
  const preview=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(preview.status).toBe('available')
  if(preview.status!=='available')return
  expect(preview.pages[0]!.content_clip.y_emu).toBe(Math.round(1.5*914400))
  expect(preview.pages[0]!.content_clip.height_emu).toBe(10058400-Math.round(1.5*914400)-914400)
 })
 it('places the authored odd header and footer on every page it paginates',()=>{
  // Two pages so `&P` differs and `&N` is the sheet's own total, not the page's.
  const {objects,geometry,part}=fixture()
  const settings={paper:'Letter' as const,orientation:'portrait' as const,scale:100,left_inches:1,right_inches:1,top_inches:10.25,bottom_inches:0.5,header_inches:0.3,footer_inches:0.25}
  objects.page_settings=[{sheet_id:'7',sheet_part:part,status:'available',warnings:['Source settings'],settings,header_footer:{odd_header:'&C&A',odd_footer:'&CPage &P of &N'}}]
  const preview=compileNativeSheetPrintPagePreviewV1([geometry],objects,{header_footer_facts:{sheet_name:'Summary',default_font_size_points:10}})
  expect(preview.status).toBe('available')
  if(preview.status!=='available')return
  expect(preview.pages.length).toBeGreaterThan(1)
  const first=preview.pages[0]!,last=preview.pages[preview.pages.length-1]!
  expect(first.header?.kind).toBe('header')
  expect(first.header?.y_css_px).toBe(0.3*96)
  expect(first.header?.x_css_px).toBe(96)
  expect(first.header?.width_css_px).toBe(816-192)
  expect(first.header?.sections).toEqual([{align:'center',runs:[{text:'Summary',font_size_points:10,bold:false,italic:false}]}])
  // The footer line ENDS at its margin measured up from the bottom edge.
  expect(first.footer?.y_css_px).toBe(1056-0.25*96)
  expect(first.footer?.sections[0]!.runs[0]!.text).toBe(`Page 1 of ${preview.pages.length}`)
  expect(last.footer?.sections[0]!.runs[0]!.text).toBe(`Page ${preview.pages.length} of ${preview.pages.length}`)
 })
 it('paints no band without the facts its codes stand for, or for a code it cannot resolve',()=>{
  const {objects,geometry,part}=fixture()
  const settings={paper:'Letter' as const,orientation:'portrait' as const,scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1,header_inches:0.3,footer_inches:0.3}
  const withHeader=(odd_header:string)=>{
   objects.page_settings=[{sheet_id:'7',sheet_part:part,status:'available',warnings:['Source settings'],settings,header_footer:{odd_header,odd_footer:''}}]
   return compileNativeSheetPrintPagePreviewV1([geometry],objects,{header_footer_facts:{sheet_name:'Summary',default_font_size_points:10}})
  }
  // The printed date is a fact no preview holds, so the whole header refuses.
  const unresolved=withHeader('&C&A &D')
  expect(unresolved.status==='available'&&unresolved.pages[0]!.header).toBeUndefined()
  // Without the worksheet name and default size there is nothing to resolve `&A` against.
  objects.page_settings=[{sheet_id:'7',sheet_part:part,status:'available',warnings:['Source settings'],settings,header_footer:{odd_header:'&C&A',odd_footer:''}}]
  const noFacts=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(noFacts.status==='available'&&noFacts.pages[0]!.header).toBeUndefined()
  // A worksheet that authors no header pair at all keeps its pages bandless.
  const {objects:plain,geometry:plainGeometry}=fixture()
  const bare=compileNativeSheetPrintPagePreviewV1([plainGeometry],plain,{header_footer_facts:{sheet_name:'Summary',default_font_size_points:10}})
  expect(bare.status==='available'&&(bare.pages[0]!.header||bare.pages[0]!.footer)).toBeFalsy()
  // Facts that are not explicit bounded plain data are a caller error, not a band.
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects,{header_footer_facts:{sheet_name:'',default_font_size_points:10}})).toThrow('explicit worksheet name')
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects,{header_footer_facts:{sheet_name:'S',default_font_size_points:0}})).toThrow('explicit worksheet name')
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

// A pageSetup that omits paperSize, orientation or scale states that
// attribute's ECMA-376 §18.3.1.63 default rather than nothing, so the page is
// produced; it is still not an authored value, and the preview says so.
describe('ECMA-376 attribute defaults inside an authored pageSetup',()=>{
 it('paginates, reports host-default origin and names the defaulted facts',()=>{
  const {objects,geometry,part}=fixture()
  const authored=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  objects.page_settings=[{
   sheet_id:'7',sheet_part:part,status:'available',warnings:['Source settings'],defaulted:['paper','scale'],
   settings:{paper:'Letter',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1},
  }]
  const preview=compileNativeSheetPrintPagePreviewV1([geometry],objects)
  expect(preview.status).toBe('available')
  if(preview.status!=='available'||authored.status!=='available')return
  expect(preview.settings_origin).toBe('host-default')
  expect(preview.warnings.some(w=>w.includes('paper, scale are ECMA-376 CT_PageSetup attribute defaults'))).toBe(true)
  expect(preview.warnings.some(w=>w.includes('host default, not authored workbook settings'))).toBe(false)
  expect(preview.pages.map(p=>[p.width_emu,p.height_emu])).toEqual(authored.pages.map(p=>[p.width_emu,p.height_emu]))
 })
 it('keeps an authored page labelled source, and rejects a malformed defaulted list',()=>{
  const {objects,geometry,part}=fixture()
  expect(compileNativeSheetPrintPagePreviewV1([geometry],objects).settings_origin).toBe('source')
  const settings={paper:'Letter' as const,orientation:'portrait' as const,scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}
  for(const defaulted of [[],['paper','paper'],['margins'],['paper','orientation','scale','paper']]){
   objects.page_settings=[{sheet_id:'7',sheet_part:part,status:'available',warnings:['Source settings'],defaulted,settings} as never]
   expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects)).toThrow('Invalid native worksheet page settings')
  }
  // Only settings this tier actually produced can carry defaulted facts.
  objects.page_settings=[{sheet_id:'7',sheet_part:part,status:'unavailable',warnings:['Unsupported'],defaulted:['paper']} as never]
  expect(()=>compileNativeSheetPrintPagePreviewV1([geometry],objects)).toThrow('Invalid native worksheet page settings')
 })
})
