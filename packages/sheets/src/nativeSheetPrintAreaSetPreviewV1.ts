import {isProjectedNativeWorkbookV2,type NativeWorkbookRenderModelV2} from './nativeRenderModelV2.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {compiledNativeSheetGeometrySourcePart,isCompiledNativeSheetGeometryV2,isCompiledNativeStoredRowSheetGeometryV1,NATIVE_SHEET_GEOMETRY_V2_LIMITS,type NativeSheetGeometryV2,type NativeSheetViewportV2} from './nativeSheetGeometryV2.js'
import {compileNativeSheetPagePreviewV1,type NativeSheetPagePreviewV1,type NativeSheetHostPagePolicyV1,type NativeSheetPagePreviewOptionsV1} from './nativeSheetPagePreviewV1.js'

function sourceAreas(source:NativeWorkbookObjectsV1,sheetId:string,part:string|undefined):readonly NativeSheetViewportV2[]{
 const entry=source.print_area_sets!==undefined?source.print_area_sets.find(s=>s.sheet_id===sheetId):source.print_areas?.find(s=>s.sheet_id===sheetId)
 if(!entry||!part||entry.sheet_part!==part)throw new TypeError('Print area set does not join the source worksheet')
 if(entry.status!=='available')throw new TypeError('Source print area set unavailable; an explicit host range is required')
 const areas='areas'in entry?entry.areas:[entry.area]
 let cells=0
 for(const area of areas){
  const rows=area.end_row-area.row+1,columns=area.end_column-area.column+1,limits=NATIVE_SHEET_GEOMETRY_V2_LIMITS
  cells+=rows*columns
  if(rows>limits.maxViewportRows||columns>limits.maxViewportColumns||cells>limits.maxViewportCells)throw new RangeError('Saved print area set exceeds native geometry viewport limits or aggregate 100000-cell budget')
 }
 return Object.freeze(areas.map(a=>Object.freeze({...a})))
}

/** Select all saved ranges in source order. Only absent additive metadata permits legacy single-area selection. */
export function selectNativeSheetPrintAreaSetV1(model:NativeWorkbookRenderModelV2,sheetId:string,objects:NativeWorkbookObjectsV1):readonly NativeSheetViewportV2[]{
 if(!isProjectedNativeWorkbookV2(model))throw new TypeError('Print area set selection requires a projected source workbook')
 const source=decodeNativeWorkbookObjectsV1(objects,model.source.package_sha256)
 return sourceAreas(source,sheetId,model.sheets.find(s=>s.id===sheetId)?.mutation_authority.source_part)
}

export interface NativeSheetPrintAreaSetPreviewV1 {
 protocol:'injoffice.xlsx.print-area-set-pages';version:1;
 areas:{area_index:number;viewport:NativeSheetViewportV2;plan:NativeSheetPagePreviewV1}[];
 total_pages:number;
}

/** All-or-nothing source-qualified planning: independent fit per range, at most 100 pages in total. */
export function compileNativeSheetPrintAreaSetPreviewV1(geometries:readonly NativeSheetGeometryV2[],objects:NativeWorkbookObjectsV1,hostPolicy?:NativeSheetHostPagePolicyV1,options?:NativeSheetPagePreviewOptionsV1):NativeSheetPrintAreaSetPreviewV1{
 const fail=():never=>{throw new TypeError('Print area set preview requires a plain array of compiled source-qualified sheet geometries')}
 if(!Array.isArray(geometries)||Object.getPrototypeOf(geometries)!==Array.prototype)return fail()
 const descriptors:Record<PropertyKey,PropertyDescriptor>=Object.getOwnPropertyDescriptors(geometries) as unknown as Record<PropertyKey,PropertyDescriptor>,length=descriptors.length?.value
 if(typeof length!=='number'||!Number.isSafeInteger(length)||length<1||length>16||Reflect.ownKeys(descriptors).length!==length+1)return fail()
 // Preserve private geometry brands, while refusing getters, holes and custom iterators.
 const owned:NativeSheetGeometryV2[]=[]
 for(let i=0;i<length;i++){
  const d=descriptors[String(i)]
  if(!d||!('value'in d)||(!isCompiledNativeSheetGeometryV2(d.value)&&!isCompiledNativeStoredRowSheetGeometryV1(d.value)))return fail()
  owned.push(d.value)
 }
 geometries=owned
 const first=geometries[0]!,part=compiledNativeSheetGeometrySourcePart(first)
 const source=decodeNativeWorkbookObjectsV1(objects,first.source_package_sha256)
 const viewports=sourceAreas(source,first.sheet_id,part)
 if(viewports.length!==geometries.length)throw new TypeError('Print area set geometry count does not match the complete source set')
 for(let i=0;i<geometries.length;i++){
  const g=geometries[i]!,v=viewports[i]!
  if(g.source_package_sha256!==first.source_package_sha256||g.document_id!==first.document_id||g.source_revision!==first.source_revision||g.sheet_id!==first.sheet_id||compiledNativeSheetGeometrySourcePart(g)!==part||g.viewport.row!==v.row||g.viewport.column!==v.column||g.viewport.end_row!==v.end_row||g.viewport.end_column!==v.end_column)throw new TypeError('Print area set geometry identity or source range order does not match')
 }
 let total_pages=0
 const areas=geometries.map((geometry,area_index)=>{
  const plan=compileNativeSheetPagePreviewV1(geometry,source,hostPolicy,options)
  total_pages+=plan.pages.length
  if(total_pages>100)throw new RangeError('Print area set preview exceeds aggregate 100-page budget')
  return {area_index,viewport:viewports[area_index]!,plan}
 })
 return {protocol:'injoffice.xlsx.print-area-set-pages',version:1,areas,total_pages}
}
