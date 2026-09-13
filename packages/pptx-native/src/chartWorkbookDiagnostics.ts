import type {NativeWorkbookV2} from '@injoffice/sheets/browser'
type NativeWorkbookUnsupportedV2=NativeWorkbookV2['unsupported'][number]

// These features cannot replace literal cell values or hidden flags. They are
// still returned as provenance. Unknown features fail closed; XLSX's mutation
// impact classification alone is not a chart-data semantic classification.
const harmless=new Set(['SHEET_PROTECTION','SHEET_VIEW_GEOMETRY','MERGED_CELLS','CONDITIONAL_FORMATTING','DATA_VALIDATION','HYPERLINKS','DRAWING_REFERENCE','TABLE_REFERENCE','RICH_SHARED_STRING'])
export function qualifyChartWorkbookDiagnostics(items:readonly NativeWorkbookUnsupportedV2[],sheetId:string,addresses:readonly string[],usesSharedStrings:boolean):readonly NativeWorkbookUnsupportedV2[]{
 if(!Array.isArray(items)||items.length>16384)throw new RangeError('workbook source diagnostic budget exceeded')
 const selected=new Set(addresses),result:NativeWorkbookUnsupportedV2[]=[]
 for(const item of items){
  if(item.scope_id!==`sheet:${sheetId}`&&item.scope_id!=='workbook'&&!item.scope_id.startsWith('style:'))continue
  if(item.cell_ref!==undefined&&!selected.has(item.cell_ref))continue
  if(item.range_ref!==undefined){
   const match=/^([A-Z]{1,3}[1-9][0-9]*)(?::([A-Z]{1,3}[1-9][0-9]*))?$/.exec(item.range_ref)
   const position=(s:string)=>{const m=/^([A-Z]+)([0-9]+)$/.exec(s)!;let c=0;for(const ch of m[1]!)c=c*26+ch.charCodeAt(0)-64;return [Number(m[2]),c] as const}
   if(!match)throw new RangeError('invalid workbook diagnostic range')
   const [r0,c0]=position(match[1]!),[r1,c1]=position(match[2]??match[1]!)
   if(r0<1||r1>1048576||c0<1||c1>16384||r1<r0||c1<c0)throw new RangeError('invalid workbook diagnostic range')
   const [a0,b0]=position(addresses[0]!),[a1,b1]=position(addresses[addresses.length-1]!)
   if(r1<a0||r0>a1||c1<b0||c0>b1)continue
  }
  if(result.length>=4096)throw new RangeError('chart source diagnostic output budget exceeded')
  const copy=Object.freeze({...item});result.push(copy)
  const style=item.capability==='styles'
  const unusedStrings=!usesSharedStrings&&item.capability==='rich-text'&&item.scope_id==='workbook'
  if(!style&&!unusedStrings&&!harmless.has(item.code))throw new RangeError(`referenced workbook source is outside the chart profile: ${item.code}`)
 }
 return Object.freeze(result)
}
