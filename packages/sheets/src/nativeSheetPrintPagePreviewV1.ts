import {
  compiledNativeSheetGeometrySourcePart,
  EMU_PER_CSS_PIXEL,
  isCompiledNativeSheetGeometryV2,
  isCompiledNativeStoredRowSheetGeometryV1,
  type NativeSheetGeometryRectV2,
  type NativeSheetGeometryV2,
} from './nativeSheetGeometryV2.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {
  compileNativeSheetPrintAreaSetPreviewV1,
  type NativeSheetPrintAreaSetPreviewV1,
} from './nativeSheetPrintAreaSetPreviewV1.js'
import type {NativeSheetHostPagePolicyV1,NativeSheetPreviewRegionV1} from './nativeSheetPagePreviewV1.js'
import type {NativeSheetCellPaintPlanV2} from './nativeSheetCellPaintV2.js'
import {
  parseNativeSheetHeaderFooterV1,
  type NativeSheetHeaderFooterSectionV1,
} from './nativeSheetHeaderFooterV1.js'

export const NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_PROTOCOL='injoffice.xlsx.print-page-preview' as const
export const NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI=96 as const

const GRID_UNCHANGED='Grid preview is unchanged.'
const NO_INVENT='Print-page preview does not invent paper, margins, scale, printer defaults or a used-range fallback.'
const APPROXIMATION='Read-only print-page preview at 96 CSS pixels per inch from source-qualified page rectangles. This is not Excel printer calibration, LibreOffice PDF raster parity, or fit-to-page qualification.'
/**
 * ECMA-376 §18.18.50 gives an absent pageSetup the application default rather
 * than no page at all, and Excel 16.112.4 paginates every such worksheet: all
 * 33 in the local hard-v2 corpus that declare margins and no pageSetup export
 * at US Letter portrait. This is that default, named as a host constant so it
 * stays a declared choice rather than a literal inside the layout, and so an
 * A4-region host can retarget it. Authored margins are always preferred; only
 * paper, orientation and scale come from here.
 */
const HOST_DEFAULT_PAGE={paper:'Letter',orientation:'portrait',scale:100} as const
const HOST_DEFAULT_DISCLOSURE='Paper, orientation and scale are a host default, not authored workbook settings: this worksheet declares page margins and no pageSetup element. Authored margins are used unchanged.'
/**
 * A pageSetup that omits an attribute states that attribute's ECMA-376
 * §18.3.1.63 default, which Excel prints; the source tier supplies it and names
 * it. That is still not an authored value, so the page reports host-default
 * origin and says which facts came from the schema rather than the workbook.
 */
const defaultedFactDisclosure=(facts:readonly string[]):string=>
 `Page ${facts.join(', ')} ${facts.length===1?'is an ECMA-376 CT_PageSetup attribute default':'are ECMA-376 CT_PageSetup attribute defaults'} supplied for an attribute this worksheet omits, not an authored workbook setting. Every authored page fact, including the margins, is used unchanged.`

export interface NativeSheetPrintPageCssRectV1 {
 readonly x_css_px:number;readonly y_css_px:number;readonly width_css_px:number;readonly height_css_px:number
}
export interface NativeSheetPrintPageRegionV1 {
 readonly kind:NativeSheetPreviewRegionV1['kind']
 readonly source_clip:NativeSheetGeometryRectV2
 readonly source_clip_css_px:NativeSheetPrintPageCssRectV1
 readonly translate_x_emu:number;readonly translate_y_emu:number
 readonly translate_x_css_px:number;readonly translate_y_css_px:number
 readonly rows:{readonly start:number;readonly end:number}
 readonly columns:{readonly start:number;readonly end:number}
}
export interface NativeSheetPrintPagePaintV1 {
 readonly geometry_sha256:string
 readonly coordinate_space:'viewport-local'
 readonly clip:NativeSheetGeometryRectV2
 readonly scale:number
 readonly translate_x_emu:number;readonly translate_y_emu:number
 readonly plan:NativeSheetCellPaintPlanV2
}
export interface NativeSheetPrintPageRasterPageV1 {
 readonly sequence:number
 readonly number:number
 readonly area_index:number
 readonly width_emu:number;readonly height_emu:number
 readonly width_css_px:number;readonly height_css_px:number
 readonly content_clip:NativeSheetGeometryRectV2
 readonly content_clip_css_px:NativeSheetPrintPageCssRectV1
 readonly source_clip:NativeSheetGeometryRectV2
 readonly source_clip_css_px:NativeSheetPrintPageCssRectV1
 readonly scale:number
 readonly translate_x_emu:number;readonly translate_y_emu:number
 readonly translate_x_css_px:number;readonly translate_y_css_px:number
 readonly rows:{readonly start:number;readonly end:number}
 readonly columns:{readonly start:number;readonly end:number}
 readonly regions?:readonly NativeSheetPrintPageRegionV1[]
 readonly paint?:NativeSheetPrintPagePaintV1
 /** Authored header and footer text for THIS page; absent when none is painted. */
 readonly header?:NativeSheetPrintPageBandV1
 readonly footer?:NativeSheetPrintPageBandV1
}
/**
 * One printed header or footer line, in page-local CSS pixels.
 *
 * ECMA-376 §18.3.1.62 measures the header and footer margins from the paper
 * edge, so the header line starts at the header margin and the footer line ENDS
 * at the footer margin measured up from the bottom; neither is an inset of the
 * body rectangle, which the preview already shortens to max(top, header) and
 * max(bottom, footer). The line spans the body's own width, which is what
 * alignWithMargins="true" (the only shape carried here) means.
 *
 * A single line per band: Excel grows the reservation for a multi-line header
 * and this tier does not reproduce that, so a `\n` in the authored text keeps
 * the band refused upstream rather than overprinting the body.
 */
export interface NativeSheetPrintPageBandV1 {
 readonly kind:'header'|'footer'
 readonly x_css_px:number;readonly width_css_px:number
 /** Top of the line box for a header; for a footer this is its bottom. */
 readonly y_css_px:number
 readonly sections:readonly NativeSheetHeaderFooterSectionV1[]
}
export interface NativeSheetPrintPagePreviewOptionsV1 {
 readonly repeat_print_titles?:true
 readonly paint_plans?:readonly NativeSheetCellPaintPlanV2[]
 /**
  * The facts the authored header and footer codes stand for. `&A` is the
  * worksheet name and an uncoded run is the workbook's Normal font at this
  * size, neither of which page geometry knows, so omitting this leaves every
  * band unpainted rather than inventing a name or a size.
  */
 readonly header_footer_facts?:{readonly sheet_name:string;readonly default_font_size_points:number}
}
type NativeSheetPrintPagePreviewBaseV1={
 readonly protocol:typeof NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_PROTOCOL
 readonly version:1
 readonly fidelity:'approximate'
 readonly read_only:true
 readonly dpi:typeof NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI
 readonly document_id:string
 readonly sheet_id:string
 readonly source_revision:string
 readonly source_package_sha256:string
 /** 'host-default' when paper, orientation or scale came from the ECMA-376 attribute default rather than the workbook. */
 readonly settings_origin:'source'|'host-default'
 readonly warnings:readonly string[]
}
export type NativeSheetPrintPagePreviewV1=NativeSheetPrintPagePreviewBaseV1&(
 |{readonly status:'available';readonly pages:readonly NativeSheetPrintPageRasterPageV1[];readonly source_plan:NativeSheetPrintAreaSetPreviewV1}
 |{readonly status:'unavailable';readonly reason:string;readonly pages:readonly []}
)

function cssPx(emu:number):number{return emu/EMU_PER_CSS_PIXEL}
function cssRect(rect:NativeSheetGeometryRectV2):NativeSheetPrintPageCssRectV1{
 return Object.freeze({x_css_px:cssPx(rect.x_emu),y_css_px:cssPx(rect.y_emu),width_css_px:cssPx(rect.width_emu),height_css_px:cssPx(rect.height_emu)})
}
function ownedArray<T>(value:unknown,length:number,accept:(item:unknown)=>item is T,message:string):T[]{
 const fail=():never=>{throw new TypeError(message)}
 if(!Array.isArray(value)||Object.getPrototypeOf(value)!==Array.prototype)return fail()
 const descriptors:Record<PropertyKey,PropertyDescriptor>=Object.getOwnPropertyDescriptors(value) as unknown as Record<PropertyKey,PropertyDescriptor>
 const n=descriptors.length?.value
 if(typeof n!=='number'||!Number.isSafeInteger(n)||n!==length||Reflect.ownKeys(descriptors).length!==n+1)return fail()
 const owned:T[]=[]
 for(let i=0;i<n;i++){
  const d=descriptors[String(i)]
  if(!d||!('value'in d)||!accept(d.value))return fail()
  owned.push(d.value)
 }
 return owned
}
function ownedGeometries(geometries:readonly NativeSheetGeometryV2[]):NativeSheetGeometryV2[]{
 const fail=():never=>{throw new TypeError('Print-page preview requires a plain array of compiled source-qualified sheet geometries')}
 if(!Array.isArray(geometries)||!Number.isSafeInteger(geometries.length)||geometries.length<1||geometries.length>16)return fail()
 return ownedArray(
  geometries,
  geometries.length,
  (item):item is NativeSheetGeometryV2=>isCompiledNativeSheetGeometryV2(item)||isCompiledNativeStoredRowSheetGeometryV1(item),
  'Print-page preview requires a plain array of compiled source-qualified sheet geometries',
 )
}
type HeaderFooterFactsV1={sheet_name:string;default_font_size_points:number}
function readOptions(options:NativeSheetPrintPagePreviewOptionsV1|undefined,geometryCount:number):{repeat?:true;paint_plans?:NativeSheetCellPaintPlanV2[];facts?:HeaderFooterFactsV1}{
 if(options===undefined)return {}
 const descriptors=Object.getOwnPropertyDescriptors(options)
 const keys=Object.keys(descriptors)
 if(!keys.length||keys.some(key=>key!=='repeat_print_titles'&&key!=='paint_plans'&&key!=='header_footer_facts')||Object.values(descriptors).some(d=>!('value'in d))){
  throw new TypeError('Print-page preview options must be explicit plain data')
 }
 let facts:HeaderFooterFactsV1|undefined
 if(Object.hasOwn(descriptors,'header_footer_facts')){
  const value=descriptors.header_footer_facts!.value as Record<string,unknown>
  const own=value&&typeof value==='object'&&!Array.isArray(value)?Object.getOwnPropertyDescriptors(value):undefined
  if(!own||Object.keys(own).length!==2||!Object.hasOwn(own,'sheet_name')||!Object.hasOwn(own,'default_font_size_points')||Object.values(own).some(d=>!('value'in d))){
   throw new TypeError('Print-page preview header and footer facts must be an explicit worksheet name and default font size')
  }
  const name=own.sheet_name!.value,points=own.default_font_size_points!.value
  if(typeof name!=='string'||!name||name.length>256||typeof points!=='number'||!Number.isFinite(points)||points<1||points>409){
   throw new TypeError('Print-page preview header and footer facts must be an explicit worksheet name and default font size')
  }
  facts={sheet_name:name,default_font_size_points:points}
 }
 if(Object.hasOwn(descriptors,'repeat_print_titles')&&descriptors.repeat_print_titles!.value!==true){
  throw new TypeError('Repeated print titles require the explicit repeat_print_titles: true option')
 }
 return {
  ...(facts?{facts}:{}),
  ...(Object.hasOwn(descriptors,'repeat_print_titles')?{repeat:true as const}:{}),
  ...(Object.hasOwn(descriptors,'paint_plans')?{
   paint_plans:ownedArray(
    descriptors.paint_plans!.value,
    geometryCount,
    (item):item is NativeSheetCellPaintPlanV2=>!!item&&typeof item==='object',
    'Print-page preview paint plans must be a plain array joining each compiled geometry',
   ),
  }:{}),
 }
}
function refuse(geometry:NativeSheetGeometryV2,reason:string,warnings:readonly string[]=[]):NativeSheetPrintPagePreviewV1{
 return Object.freeze({
  protocol:NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_PROTOCOL,version:1,fidelity:'approximate' as const,read_only:true as const,
  dpi:NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI,document_id:geometry.document_id,sheet_id:geometry.sheet_id,
  source_revision:geometry.source_revision,source_package_sha256:geometry.source_package_sha256,settings_origin:'source' as const,
  status:'unavailable' as const,reason,pages:Object.freeze([]) as readonly [],
  warnings:Object.freeze([...warnings,APPROXIMATION,NO_INVENT,GRID_UNCHANGED]),
 })
}
/** Either a refusal, or the host policy to paginate with (absent when the source authored its own). */
type PageSettingsResolutionV1={readonly reason:string}|{
 readonly host_policy:NativeSheetHostPagePolicyV1|undefined
 /** Source-tier settings whose value is a schema default rather than authored. */
 readonly defaulted?:readonly string[]
}
function resolvePageSettings(source:NativeWorkbookObjectsV1,sheetId:string,part:string|undefined):PageSettingsResolutionV1{
 if(source.page_settings===undefined)return {reason:`Source page settings are missing. ${NO_INVENT} ${GRID_UNCHANGED}`}
 const candidates=source.page_settings.filter(s=>s.sheet_id===sheetId)
 if(candidates.length!==1)return {reason:`Source page settings do not join the source worksheet. ${NO_INVENT} ${GRID_UNCHANGED}`}
 const page=candidates[0]!
 if(!part||page.sheet_part!==part)return {reason:`Source page settings do not join the source worksheet part. ${NO_INVENT} ${GRID_UNCHANGED}`}
 // No pageSetup element is authored, so no paper is authored to contradict. Use
 // the host default for paper, orientation and scale, and the worksheet's own
 // margins unchanged. A pageSetup that exists but is ambiguous or unsupported is
 // a different fact and still refuses: that would be overriding an authored value.
 if(page.status==='margins-only'){
  if(!page.margins)return {reason:`Source page margins are missing. ${NO_INVENT} ${GRID_UNCHANGED}`}
  return {host_policy:{
   kind:'explicit-host-page-policy-v1',
   paper:HOST_DEFAULT_PAGE.paper,orientation:HOST_DEFAULT_PAGE.orientation,scale:HOST_DEFAULT_PAGE.scale,
   left_inches:page.margins.left_inches,right_inches:page.margins.right_inches,
   top_inches:page.margins.top_inches,bottom_inches:page.margins.bottom_inches,
   // Authored header/footer margins are the worksheet's own too, and they
   // reserve body height; dropping them would print a taller body than Excel.
   ...(page.margins.header_inches!==undefined?{header_inches:page.margins.header_inches}:{}),
   ...(page.margins.footer_inches!==undefined?{footer_inches:page.margins.footer_inches}:{}),
  }}
 }
 if(page.status!=='available'||!page.settings){
  const detail=page.warnings[0]??'Source page settings are printer-dependent or unsupported'
  return {reason:`${detail} ${NO_INVENT} ${GRID_UNCHANGED}`}
 }
 if(page.settings.fit_to_page)return {reason:`Source page settings include fit-to-page. Print-page preview does not invent a fit scale and is not Excel fit-to-page qualification. ${GRID_UNCHANGED}`}
 return {host_policy:undefined,...(page.defaulted?{defaulted:page.defaulted}:{})}
}
function printAreaReason(source:NativeWorkbookObjectsV1,sheetId:string,part:string|undefined):string|undefined{
 const entry=source.print_area_sets!==undefined?source.print_area_sets.find(s=>s.sheet_id===sheetId):source.print_areas?.find(s=>s.sheet_id===sheetId)
 if(!entry||!part||entry.sheet_part!==part)return `Saved print area does not join the source worksheet. Print-page preview does not invent a used-range or A1 fallback. ${GRID_UNCHANGED}`
 if(entry.status!=='available'){
  const detail=entry.warnings[0]??'Source print area unavailable'
  return `${detail} Print-page preview does not invent a used-range or A1 fallback. ${GRID_UNCHANGED}`
 }
 return undefined
}
function rasterPage(
 page:{number:number;width_emu:number;height_emu:number;content_clip:NativeSheetGeometryRectV2;source_clip:NativeSheetGeometryRectV2;scale:number;translate_x_emu:number;translate_y_emu:number;rows:{start:number;end:number};columns:{start:number;end:number};regions?:NativeSheetPreviewRegionV1[]},
 sequence:number,areaIndex:number,paint?:NativeSheetCellPaintPlanV2,
 bands?:{header?:NativeSheetPrintPageBandV1;footer?:NativeSheetPrintPageBandV1},
):NativeSheetPrintPageRasterPageV1{
 const regions=page.regions?.map(region=>Object.freeze({
  kind:region.kind,source_clip:region.source_clip,source_clip_css_px:cssRect(region.source_clip),
  translate_x_emu:region.translate_x_emu,translate_y_emu:region.translate_y_emu,
  translate_x_css_px:cssPx(region.translate_x_emu),translate_y_css_px:cssPx(region.translate_y_emu),
  rows:region.rows,columns:region.columns,
 }))
 return Object.freeze({
  sequence,number:page.number,area_index:areaIndex,
  width_emu:page.width_emu,height_emu:page.height_emu,
  width_css_px:cssPx(page.width_emu),height_css_px:cssPx(page.height_emu),
  content_clip:page.content_clip,content_clip_css_px:cssRect(page.content_clip),
  source_clip:page.source_clip,source_clip_css_px:cssRect(page.source_clip),
  scale:page.scale,translate_x_emu:page.translate_x_emu,translate_y_emu:page.translate_y_emu,
  translate_x_css_px:cssPx(page.translate_x_emu),translate_y_css_px:cssPx(page.translate_y_emu),
  rows:page.rows,columns:page.columns,
  ...(regions?{regions:Object.freeze(regions)}:{}),
  ...(bands?.header?{header:bands.header}:{}),
  ...(bands?.footer?{footer:bands.footer}:{}),
  ...(paint?{paint:Object.freeze({
   geometry_sha256:paint.geometry_sha256,coordinate_space:'viewport-local' as const,clip:page.source_clip,
   scale:page.scale,translate_x_emu:page.translate_x_emu,translate_y_emu:page.translate_y_emu,plan:paint,
  })}:{}),
 })
}

const CSS_PX_PER_INCH=NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI
/**
 * The printed header and footer lines for one page, or nothing.
 *
 * Nothing is the answer whenever any input the printed result depends on is
 * missing or unsupported: no authored pair, no worksheet name and default size
 * to resolve `&A` and an uncoded run against, no authored header or footer
 * margin to place the line at, or a code this tier does not resolve. A band is
 * also dropped when its own margin would put it inside the body rectangle,
 * which is the overlong-header case the preview already declines to reproduce.
 */
function headerFooterBands(
 authored:{odd_header:string;odd_footer:string}|undefined,
 margins:{header_inches?:number;footer_inches?:number}|undefined,
 facts:HeaderFooterFactsV1|undefined,
 page:{width_emu:number;height_emu:number;number:number;content_clip:NativeSheetGeometryRectV2},
 sequence:number,total:number,
):{header?:NativeSheetPrintPageBandV1;footer?:NativeSheetPrintPageBandV1}|undefined{
 if(!authored||!facts||!margins)return undefined
 const x=cssPx(page.content_clip.x_emu),width=cssPx(page.content_clip.width_emu)
 const bodyTop=cssPx(page.content_clip.y_emu),bodyBottom=bodyTop+cssPx(page.content_clip.height_emu)
 const height=cssPx(page.height_emu)
 const band=(kind:'header'|'footer',code:string,inches:number|undefined):NativeSheetPrintPageBandV1|undefined=>{
  if(!code||inches===undefined)return undefined
  const y=kind==='header'?inches*CSS_PX_PER_INCH:height-inches*CSS_PX_PER_INCH
  if(kind==='header'?y>=bodyTop:y<=bodyBottom)return undefined
  const sections=parseNativeSheetHeaderFooterV1(code,{
   sheet_name:facts.sheet_name,page_number:sequence,page_count:total,
   default_font_size_points:facts.default_font_size_points,
  })
  if(!sections||sections.some(section=>section.runs.some(run=>/[\r\n]/.test(run.text))))return undefined
  return Object.freeze({kind,x_css_px:x,width_css_px:width,y_css_px:y,sections})
 }
 const header=band('header',authored.odd_header,margins.header_inches)
 const footer=band('footer',authored.odd_footer,margins.footer_inches)
 return header||footer?{...(header?{header}:{}),...(footer?{footer}:{})}:undefined
}

/** Source-only print-page rectangles at 96 CSS px/in. Missing or printer-dependent
 * setup, fit-to-page, and missing saved print areas refuse without host fallbacks.
 */
export function compileNativeSheetPrintPagePreviewV1(
 geometries:readonly NativeSheetGeometryV2[],
 objects:NativeWorkbookObjectsV1,
 options?:NativeSheetPrintPagePreviewOptionsV1,
):NativeSheetPrintPagePreviewV1{
 const owned=ownedGeometries(geometries)
 const first=owned[0]!
 const parsed=readOptions(options,owned.length)
 const source=decodeNativeWorkbookObjectsV1(objects,first.source_package_sha256)
 const part=compiledNativeSheetGeometrySourcePart(first)
 const settings=resolvePageSettings(source,first.sheet_id,part)
 if('reason'in settings)return refuse(first,settings.reason,source.page_settings?.find(s=>s.sheet_id===first.sheet_id)?.warnings??[])
 const hostPolicy=settings.host_policy
 const defaulted=settings.defaulted
 const areaRefusal=printAreaReason(source,first.sheet_id,part)
 if(areaRefusal)return refuse(first,areaRefusal)
 if(parsed.paint_plans){
  for(let i=0;i<owned.length;i++){
   const plan=parsed.paint_plans[i]!,geometry=owned[i]!
   if(plan.geometry_sha256!==geometry.geometry_sha256||plan.sheet_id!==geometry.sheet_id||plan.source_package_sha256!==geometry.source_package_sha256){
    throw new TypeError('Print-page preview paint plans must join compiled geometry identity')
   }
  }
 }
 let source_plan:NativeSheetPrintAreaSetPreviewV1
 try{
  source_plan=compileNativeSheetPrintAreaSetPreviewV1(owned,source,hostPolicy,parsed.repeat?{repeat_print_titles:true}:undefined)
 }catch(error){
  if(error instanceof RangeError)return refuse(first,`${error.message} ${GRID_UNCHANGED}`)
  if(error instanceof TypeError){
   const message=error.message
   if(message.includes('plain array')||message.includes('geometry identity')||message.includes('geometry count')||message.includes('compiled source'))throw error
   return refuse(first,`${message} ${GRID_UNCHANGED}`)
  }
  throw error
 }
 const pages:NativeSheetPrintPageRasterPageV1[]=[]
 // `&P` and `&N` count across everything this sheet prints, so the total is
 // known before the first page is built, not after it.
 const total=source_plan.areas.reduce((sum,area)=>sum+area.plan.pages.length,0)
 const entry=source.page_settings!.find(s=>s.sheet_id===first.sheet_id)!
 const bandMargins=hostPolicy??entry.settings
 for(const area of source_plan.areas){
  const paint=parsed.paint_plans?.[area.area_index]
  for(const page of area.plan.pages){
   const sequence=pages.length+1
   pages.push(rasterPage(page,sequence,area.area_index,paint,headerFooterBands(entry.header_footer,bandMargins,parsed.facts,page,sequence,total)))
  }
 }
 return Object.freeze({
  protocol:NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_PROTOCOL,version:1,fidelity:'approximate' as const,read_only:true as const,
  dpi:NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI,document_id:first.document_id,sheet_id:first.sheet_id,
  source_revision:first.source_revision,source_package_sha256:first.source_package_sha256,
  settings_origin:hostPolicy||defaulted?'host-default' as const:'source' as const,
  status:'available' as const,pages:Object.freeze(pages),source_plan,
  warnings:Object.freeze([
   ...new Set(source_plan.areas.flatMap(area=>area.plan.warnings)),
   ...(hostPolicy?[HOST_DEFAULT_DISCLOSURE]:defaulted?[defaultedFactDisclosure(defaulted)]:[]),
   APPROXIMATION,
   'Hosts map viewport-local paint with x * scale + translate, clip to each page source_clip, and raster isolated pages at 96 CSS pixels per inch. Fractional A4 CSS sizes are retained; paper size is not rounded.',
  ]),
 })
}
