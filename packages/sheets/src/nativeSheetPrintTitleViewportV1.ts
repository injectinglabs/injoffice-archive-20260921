import {isProjectedNativeWorkbookV2,type NativeWorkbookRenderModelV2} from './nativeRenderModelV2.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {snapshotNativeSheetViewportV2,type NativeSheetViewportV2} from './nativeSheetGeometryV2.js'
import type {NativeSheetPrintTitlesV1} from './nativeSheetPrintTitlesV1.js'

/** Internal rectangle union. The gap supplies geometry only, never extra print content. */
export function expandedNativeSheetPrintTitleViewportV1(viewport:NativeSheetViewportV2,titles:NativeSheetPrintTitlesV1):NativeSheetViewportV2 {
 const body=snapshotNativeSheetViewportV2(viewport)
 if(titles.status!=='available')throw new TypeError('Saved print titles unavailable')
 return Object.freeze(snapshotNativeSheetViewportV2({
  row:Math.min(body.row,titles.rows?.start??body.row),
  column:Math.min(body.column,titles.columns?.start??body.column),
  end_row:Math.max(body.end_row,titles.rows?.end??body.end_row),
  end_column:Math.max(body.end_column,titles.columns?.end??body.end_column),
 }))
}

/** Source-bound geometry selection for a body range plus saved heading rows/columns.
 * Compile this viewport, then pass the original body_viewport to page planning.
 * This does not authorize printing intervening cells or alter the saved print area.
 */
export function selectNativeSheetPrintTitleViewportV1(model:NativeWorkbookRenderModelV2,sheetId:string,viewport:NativeSheetViewportV2,objects:NativeWorkbookObjectsV1):NativeSheetViewportV2 {
 if(!isProjectedNativeWorkbookV2(model))throw new TypeError('Print heading selection requires a projected source workbook')
 const source=decodeNativeWorkbookObjectsV1(objects,model.source.package_sha256)
 const part=model.sheets.find(s=>s.id===sheetId)?.mutation_authority.source_part
 const titles=source.print_titles?.find(s=>s.sheet_id===sheetId)
 if(!part||!titles||titles.sheet_part!==part||titles.status!=='available')throw new TypeError('Saved print titles unavailable or do not join the source worksheet part')
 return expandedNativeSheetPrintTitleViewportV1(viewport,titles)
}
