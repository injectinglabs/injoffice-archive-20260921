import type {NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import type {NativeWorkbookFillV2} from './nativeContractV2.generated.js'

/** Read-only table fill decoration. Coordinates are zero-based. No mutation authority. */
export function nativeTableFillPreview(objects:NativeWorkbookObjectsV1,revision:string,sheetPart:string,row:number,column:number,fill:NativeWorkbookFillV2|undefined,styleID:number):string|undefined {
 if(objects.package_sha256!==revision||!Number.isSafeInteger(row)||!Number.isSafeInteger(column)||row<0||column<0)return undefined
 // The engine qualifies palette emission only when source fill record zero is
 // a supported no-fill record. Other records and absent provenance stay intact.
 if(!fill||fill.color!==undefined||!(fill.origin==='implicit-default'||fill.origin==='styles-record'&&fill.fill_id===0))return undefined
 const matches=objects.tables.flatMap(table=>{
  if(table.sheet_part!==sheetPart)return []
  const range=/^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(table.ref);if(!range)return []
  const col=(s:string)=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1
  const left=col(range[1]!),right=col(range[3]!),top=Number(range[2])-1,bottom=Number(range[4])-1
  if(left>right||top>bottom||right>=16384||bottom>=1048576||row<top||row>bottom||column<left||column>right)return []
  return [{table,top,bottom}]
 })
 if(matches.length!==1)return undefined
 const {table,top,bottom}=matches[0]!,palette=table.fill_preview
 if(!palette||!palette.fill_style_ids.includes(styleID)||row>bottom-table.total_rows)return undefined
 if(table.header_rows&&row===top)return palette.header
 return table.row_stripes&&(row-top-table.header_rows)%2===0?palette.stripe:palette.body
}

/** White bold header text is qualified only for source default-font styles. */
export function nativeTableHeaderTextPreview(objects:NativeWorkbookObjectsV1,revision:string,sheetPart:string,row:number,column:number,fill:NativeWorkbookFillV2|undefined,styleID:number):boolean {
 const color=nativeTableFillPreview(objects,revision,sheetPart,row,column,fill,styleID)
 if(!color)return false
 return objects.tables.some(table=>{const range=/^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/.exec(table.ref);if(!range)return false;const col=(s:string)=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;return table.sheet_part===sheetPart&&column>=col(range[1]!)&&column<=col(range[3]!)&&table.header_rows===1&&table.fill_preview?.header_font_style_ids.includes(styleID)&&Number(range[2])-1===row&&table.fill_preview.header===color})
}
