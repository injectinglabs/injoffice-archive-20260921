import type {NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
export type NativeTableBorderSideV1={style:'solid'|'double'|'none';color:string;widthPoints:number}
export type NativeTableBorderEdgesV1=Partial<Record<'top'|'right'|'bottom'|'left',NativeTableBorderSideV1>>
function range(ref:string){const m=/^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);if(!m)return;const col=(s:string)=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;const r={left:col(m[1]!),right:col(m[3]!),top:Number(m[2])-1,bottom:Number(m[4])-1};if(r.left>r.right||r.top>r.bottom||r.right>=16384||r.bottom>=1048576)return;return r}
const inside=(r:NonNullable<ReturnType<typeof range>>,row:number,column:number)=>row>=r.top&&row<=r.bottom&&column>=r.left&&column<=r.right

/** No guessed shared-edge precedence: adjacent explicit/unknown borders suppress
 * the table edge. Null neighbors denote only actual worksheet boundaries. */
export function nativeTableBorderPreview(objects:NativeWorkbookObjectsV1,revision:string,sheetPart:string,row:number,column:number,styleID:number,neighbors:Record<'top'|'right'|'bottom'|'left',number|null|undefined>):NativeTableBorderEdgesV1|undefined{
 if(objects.package_sha256!==revision||!Number.isSafeInteger(row)||!Number.isSafeInteger(column))return
 const tables=objects.tables.flatMap(table=>{const bounds=range(table.ref);return table.sheet_part===sheetPart&&bounds?[{table,bounds}]:[]})
 const matches=tables.filter(t=>inside(t.bounds,row,column));if(matches.length!==1)return
 const {table,bounds}=matches[0]!,border=table.border_preview;if(!border||!border.style_ids.includes(styleID))return
 const result:NativeTableBorderEdgesV1={}
 for(const [side,dr,dc] of [['top',-1,0],['right',0,1],['bottom',1,0],['left',0,-1]] as const){
  const nr=row+dr,nc=column+dc,neighbor=neighbors[side],outside=nr<0||nc<0||nr>=1048576||nc>=16384
  if(outside?neighbor!==null:neighbor===null||neighbor===undefined||!border.style_ids.includes(neighbor))continue
  if(tables.some(t=>t.table.part!==table.part&&inside(t.bounds,nr,nc)))continue
  const internalVertical=side==='left'&&column>bounds.left||side==='right'&&column<bounds.right
  const totals=table.total_rows===1&&(side==='top'&&row===bounds.bottom||side==='bottom'&&row===bounds.bottom-1)
  result[side]={style:internalVertical?'none':totals?'double':'solid',color:totals?border.totals_color:border.color,widthPoints:internalVertical?0:totals?border.totals_width_points:border.width_points}
 }
 return result
}
