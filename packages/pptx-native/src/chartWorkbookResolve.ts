import type {NativeWorkbookV2,NativeWorkbookCellV2,NativeWorkbookSheetV2} from '@injoffice/sheets/browser'
import {admitChartWorkbookValues} from './chartWorkbookValueAdmission.js'
import {nativeChartDecimal} from './chartDecimalValidation.js'
import {qualifyChartWorkbookDiagnostics} from './chartWorkbookDiagnostics.js'
import {parseChartWorkbookRange,chartWorkbookCellAddress} from './chartWorkbookRange.js'
import type {ChartWorkbookBinding,ChartWorkbookReference,ChartWorkbookResolvedValues,ChartWorkbookCellVisibility} from './chartWorkbookTypes.js'

const hex=/^[a-f0-9]{64}$/
function part(value:string):boolean{return typeof value==='string'&&value.length>0&&value.length<=1024&&!value.startsWith('/')&&!/[\\?#\p{Cc}]/u.test(value)&&value.split('/').every(s=>s!==''&&s!=='.'&&s!=='..')}

/** Pure join over a validated XLSX V2 extractor result. This does not interpret
 * worksheet XML, recalculate formulas, or read formula/chart cached values.
 * Its caller must bind the injected extractor to the exact embedded bytes.
 * Initial chart integration requires explicit plotVisOnly=false. Visibility
 * source metadata is retained without applying an invented filtering policy. */
export function resolveChartWorkbookReferences(binding:ChartWorkbookBinding,references:readonly ChartWorkbookReference[],workbook:NativeWorkbookV2):readonly ChartWorkbookResolvedValues[]{
 if(typeof binding.relationshipId!=='string'||binding.relationshipId.length<1||binding.relationshipId.length>1024||/[\p{Cc}]/u.test(binding.relationshipId)||binding.autoUpdate!==undefined&&typeof binding.autoUpdate!=='boolean'||hex.exec(binding.sha256)?.[0]!==binding.sha256||!part(binding.part)||!Number.isSafeInteger(binding.byteLength)||binding.byteLength<1||binding.byteLength>8*1024*1024||references.length<1||references.length>64)throw new RangeError('invalid chart workbook binding or reference budget')
 if(workbook.protocol!=='injoffice.xlsx.native'||workbook.version!==2||workbook.source.authority!=='exact-package-bytes'||workbook.source.package_sha256!==`sha256:${binding.sha256}`||workbook.revision!==`rev:${binding.sha256}`||!part(workbook.source.workbook_part)||!Array.isArray(workbook.sheets)||workbook.sheets.length>1024)throw new RangeError('XLSX source identity does not match the embedded workbook')
 let totalPoints=0,totalScanned=0,totalText=0,totalDiagnostics=0
 const indices=new Map<NativeWorkbookSheetV2,{cells:Map<number,NativeWorkbookCellV2>;rows:Map<number,boolean>;columns:Map<number,boolean>}>()
 const indexed=(sheet:NativeWorkbookSheetV2)=>{
  const old=indices.get(sheet);if(old)return old
  if(!Array.isArray(sheet.cells)||sheet.cells.length>1000000||!Array.isArray(sheet.rows)||sheet.rows.length>1000000||!Array.isArray(sheet.columns)||sheet.columns.length>16384||totalScanned+sheet.cells.length+sheet.rows.length+sheet.columns.length>2000000||!part(sheet.part_name))throw new RangeError('workbook source-record scan budget exceeded')
  totalScanned+=sheet.cells.length+sheet.rows.length+sheet.columns.length
  const index=new Map<number,NativeWorkbookCellV2>();let previous=-1
  for(const cell of sheet.cells){
   if(!Number.isInteger(cell.row)||cell.row<0||cell.row>=1048576||!Number.isInteger(cell.column)||cell.column<0||cell.column>=16384)throw new RangeError('invalid workbook cell coordinates')
   const position=cell.row*16384+cell.column
   if(position<=previous)throw new RangeError('workbook cells are not unique source order')
   previous=position;index.set(position,cell)
  }
  const rows=new Map<number,boolean>(),columns=new Map<number,boolean>()
  for(const row of sheet.rows){
   if(!Number.isInteger(row.row)||row.row<0||row.row>=1048576||rows.has(row.row)||typeof row.hidden!=='boolean')throw new RangeError('invalid or duplicate workbook row visibility')
   rows.set(row.row,row.hidden)
  }
  for(const column of sheet.columns){
   if(!Number.isInteger(column.column)||!Number.isInteger(column.end_column)||column.column<0||column.end_column<column.column||column.end_column>=16384||typeof column.hidden!=='boolean')throw new RangeError('invalid workbook column visibility')
   for(let i=column.column;i<=column.end_column;i++){if(columns.has(i))throw new RangeError('ambiguous workbook column visibility');columns.set(i,column.hidden)}
  }
  const result={cells:index,rows,columns};indices.set(sheet,result);return result
 }
 return Object.freeze(references.map(reference=>{
  if(reference.kind!=='numRef'&&reference.kind!=='strRef'||typeof reference.cachePresent!=='boolean')throw new RangeError('invalid chart reference kind')
  const range=parseChartWorkbookRange(reference.formula),r=reference.range
  if(!r||r.sheet!==range.sheet||r.startRow!==range.startRow||r.startColumn!==range.startColumn||r.endRow!==range.endRow||r.endColumn!==range.endColumn||r.count!==range.count)throw new RangeError('chart reference coordinates do not match the source formula')
  totalPoints+=range.count;if(totalPoints>16384)throw new RangeError('chart workbook point budget exceeded')
  // Exact name matching avoids introducing a second worksheet-name case policy.
  const matches=workbook.sheets.filter(s=>s.name===range.sheet)
  if(matches.length!==1)throw new RangeError('chart reference worksheet is missing or ambiguous')
  const sheet=matches[0]!,cells=indexed(sheet),values:string[]=[],addresses:string[]=[],visibility:ChartWorkbookCellVisibility[]=[]
  if(!['visible','hidden','veryHidden'].includes(sheet.state)||!Array.isArray(sheet.rows)||sheet.rows.length>1000000||!Array.isArray(sheet.columns)||sheet.columns.length>16384)throw new RangeError('invalid worksheet visibility metadata')
  for(let i=0;i<range.count;i++){
   const row=range.startRow+(range.endRow!==range.startRow?i:0),column=range.startColumn+(range.endColumn!==range.startColumn?i:0),address=chartWorkbookCellAddress(row,column),cell=cells.cells.get(row*16384+column)
   if(!cell||cell.ref!==address||cell.formula!==undefined||cell.value===undefined)throw new RangeError('referenced cell is missing, formula-backed, or not a literal source value')
   const value=cell.value;let text:string
   if(reference.kind==='numRef'){
    if(value.kind!=='number'||value.storage!=='number'||value.rich||value.runs!==undefined||value.text!==undefined||typeof value.lexical!=='string'||!nativeChartDecimal(value.lexical))throw new RangeError('numeric chart reference requires a bounded literal numeric cell')
    text=value.lexical
   }else{
    if(value.kind!=='string'||!['shared','inline'].includes(value.storage)||value.rich||value.runs!==undefined||typeof value.text!=='string')throw new RangeError('category chart reference requires a plain shared or inline string cell')
    text=value.text;totalText+=text.length;if(totalText>32768)throw new RangeError('chart workbook category text budget exceeded')
   }
   const rowHidden=cells.rows.get(row),columnHidden=cells.columns.get(column),defaultRowsHidden=sheet.sheet_format?.zero_height
   for(const flag of [rowHidden,columnHidden,defaultRowsHidden])if(flag!==undefined&&typeof flag!=='boolean')throw new RangeError('invalid workbook hidden flag')
   visibility.push(Object.freeze({sheetState:sheet.state,...(rowHidden===undefined?{}:{rowHidden}),...(columnHidden===undefined?{}:{columnHidden}),...(defaultRowsHidden===undefined?{}:{defaultRowsHidden})}))
   values.push(text);addresses.push(address)
  }
  const sourceDiagnostics=qualifyChartWorkbookDiagnostics(workbook.unsupported,sheet.id,addresses,reference.kind==='strRef')
  totalDiagnostics+=sourceDiagnostics.length;if(totalDiagnostics>4096)throw new RangeError('chart source diagnostic output budget exceeded')
  return admitChartWorkbookValues(Object.freeze({dataOrigin:'embedded-workbook' as const,formula:reference.formula,kind:reference.kind,workbookPart:binding.part,workbookRelationshipId:binding.relationshipId,workbookSHA256:binding.sha256,workbookRevision:workbook.revision,sheetId:sheet.id,sheetName:sheet.name,sheetPart:sheet.part_name,addresses:Object.freeze(addresses),values:Object.freeze(values),chartCacheIgnored:reference.cachePresent,visibility:Object.freeze(visibility),sourceDiagnostics}))
 }))
}
