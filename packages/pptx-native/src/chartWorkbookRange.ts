/** The original formula spelling is retained separately from its exact cell
 * coordinates. This grammar accepts one sheet-qualified A1 row/column range,
 * never an expression, external link, defined name or chart cache. */
export interface ChartWorkbookRange {
 readonly sheet:string
 readonly startRow:number
 readonly startColumn:number
 readonly endRow:number
 readonly endColumn:number
 readonly count:number
}
const pattern=/^(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$/
export function parseChartWorkbookRange(formula:string):ChartWorkbookRange{
 if(typeof formula!=='string'||new TextEncoder().encode(formula).length>1024)throw new RangeError('chart reference formula budget exceeded')
 const match=pattern.exec(formula)
 if(!match||match[0]!==formula)throw new RangeError('chart data requires a direct sheet-qualified A1 range')
 const sheet=match[1]===undefined?match[2]!:match[1].replaceAll("''","'")
 if(!sheet||new TextEncoder().encode(sheet).length>256||/[\p{Cc}\p{Cs}\[\]:*?/\\]/u.test(sheet)||sheet.startsWith("'")||sheet.endsWith("'"))throw new RangeError('invalid chart reference sheet name')
 const column=(raw:string)=>{let n=0;for(const c of raw.toUpperCase())n=n*26+c.charCodeAt(0)-64;if(n<1||n>16384)throw new RangeError('chart reference column outside worksheet');return n-1}
 const row=(raw:string)=>{const n=Number(raw);if(n<1||n>1048576)throw new RangeError('chart reference row outside worksheet');return n-1}
 const startRow=row(match[4]!),startColumn=column(match[3]!),endRow=match[6]===undefined?startRow:row(match[6]),endColumn=match[5]===undefined?startColumn:column(match[5])
 if(endRow<startRow||endColumn<startColumn||endRow!==startRow&&endColumn!==startColumn)throw new RangeError('chart reference must be an increasing one-dimensional range')
 const count=(endRow-startRow+1)*(endColumn-startColumn+1)
 if(count>256)throw new RangeError('chart reference point budget exceeded')
 return {sheet,startRow,startColumn,endRow,endColumn,count}
}
export function chartWorkbookCellAddress(row:number,column:number):string{
 if(!Number.isInteger(row)||row<0||row>=1048576||!Number.isInteger(column)||column<0||column>=16384)throw new RangeError('invalid chart cell coordinate')
 let letters='',n=column+1
 while(n>0){n--;letters=String.fromCharCode(65+n%26)+letters;n=Math.floor(n/26)}
 return letters+String(row+1)
}
