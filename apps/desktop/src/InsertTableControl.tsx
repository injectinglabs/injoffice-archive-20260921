import {useState} from 'react'
export default function InsertTableControl({disabled,onInsert}:{disabled:boolean;onInsert(rows:number,columns:number):void}) {
  const [open,setOpen]=useState(false),[rows,setRows]=useState('3'),[columns,setColumns]=useState('3')
  const r=Number(rows),c=Number(columns),valid=Number.isInteger(r)&&r>=1&&r<=20&&Number.isInteger(c)&&c>=1&&c<=12
  return <div className="document-insert-table"><button disabled={disabled} aria-expanded={open} onClick={()=>setOpen(value=>!value)}>Insert table</button>{open&&<fieldset><legend>Table size</legend><label>Rows<input aria-label="Table rows" type="number" min="1" max="20" value={rows} onChange={event=>setRows(event.target.value)} /></label><label>Columns<input aria-label="Table columns" type="number" min="1" max="12" value={columns} onChange={event=>setColumns(event.target.value)} /></label><button disabled={disabled||!valid} onClick={()=>{onInsert(r,c);setOpen(false)}}>Insert</button><button onClick={()=>setOpen(false)}>Cancel</button></fieldset>}</div>
}
