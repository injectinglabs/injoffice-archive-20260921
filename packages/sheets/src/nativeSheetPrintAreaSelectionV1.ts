import {isProjectedNativeWorkbookV2,type NativeWorkbookRenderModelV2} from './nativeRenderModelV2.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {NATIVE_SHEET_GEOMETRY_V2_LIMITS,type NativeSheetViewportV2} from './nativeSheetGeometryV2.js'

/** Select the complete saved rectangle, independently of page-settings policy.
 * Missing, unsupported or oversized saved areas require an explicit host range;
 * this helper never truncates a saved area or silently falls back to A1.
 */
export function selectNativeSheetPrintAreaV1(model:NativeWorkbookRenderModelV2,sheetId:string,objects:NativeWorkbookObjectsV1):NativeSheetViewportV2{
 if(!isProjectedNativeWorkbookV2(model))throw new TypeError('Print area selection requires a projected source workbook')
 const source=decodeNativeWorkbookObjectsV1(objects,model.source.package_sha256)
 const sheet=model.sheets.find(s=>s.id===sheetId)
 const entry=source.print_areas?.find(s=>s.sheet_id===sheetId)
 if(!sheet||!entry||entry.sheet_part!==sheet.mutation_authority.source_part)throw new TypeError('Print area does not join the source worksheet')
 if(entry.status!=='available')throw new TypeError('Source print area unavailable; an explicit host range is required')
 const rows=entry.area.end_row-entry.area.row+1,columns=entry.area.end_column-entry.area.column+1
 const limits=NATIVE_SHEET_GEOMETRY_V2_LIMITS
 if(rows>limits.maxViewportRows||columns>limits.maxViewportColumns||rows*columns>limits.maxViewportCells)throw new RangeError('Saved print area exceeds native geometry viewport limits')
 return Object.freeze({...entry.area})
}
