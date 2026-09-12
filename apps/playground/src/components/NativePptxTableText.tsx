import {useEffect,useRef,useState} from 'react'
import {createPptxWasmClient,type NativePptxTableInspection} from '@injoffice/pptx-wasm'
import {DsButton} from '../design-system/primitives'
import './native-pptx-table-text.css'

export function NativePptxTableTextResult({inspection}:{inspection:NativePptxTableInspection}){
 return <article aria-label="Read-only presentation table text" className="pptx-table-text-result">
  <p role="status">{inspection.tables.length} {inspection.tables.length===1?'table':'tables'} inspected; {inspection.omissions.length} {inspection.omissions.length===1?'omission':'omissions'}. The original presentation and editing permissions are unchanged.</p>
  <p className="ds-muted">This is source text in reading order, not PowerPoint rendering. The table layout below is for reading; authored borders, fills, fonts and spacing are not reproduced.</p>
  {inspection.tables.map((table,index)=>{
   const rows=Array.from(new Set(table.cells.map(cell=>cell.row)))
   return <section key={`${table.slide_id}:${table.object_id}`} aria-label={`Source table ${index+1} on slide ${table.slide_index+1}`}>
    <h4>Slide {table.slide_index+1}, table {index+1}</h4>
    <div className="pptx-table-text-scroll"><table className="ds-table"><caption className="visually-hidden">Source cell text; display styling is not authored PowerPoint styling</caption><tbody>{rows.map(row=><tr key={row}>{table.cells.filter(cell=>cell.row===row).map(cell=><td key={cell.column} data-source-cell={`${row+1}:${cell.column+1}`}>{cell.paragraphs.map((text,at)=><p key={at} dir="auto">{text||<span className="ds-muted">[Empty paragraph]</span>}</p>)}</td>)}</tr>)}</tbody></table></div>
    <details><summary>Source geometry and inspection limits</summary>
     <p>Object {table.object_id}. Stored frame in EMU: x {table.rect.x}, y {table.rect.y}, width {table.rect.width}, height {table.rect.height}. These measurements are not applied to this text view.</p>
     <ul>{table.warnings.map((warning,at)=><li key={at}>{warning}</li>)}</ul>
    </details>
   </section>
  })}
  {!inspection.tables.length&&<p>No table text matches this inspection profile. The existing presentation preview and its coverage warnings remain available.</p>}
  {!!inspection.omissions.length&&<details open><summary>Omitted source content ({inspection.omissions.length})</summary><ul>{inspection.omissions.map((omission,index)=><li key={index}>Slide {omission.slide_id}, object {omission.object_id}: {omission.reason}</li>)}</ul></details>}
 </article>
}

/** Separate browser-only operation; never uploads even in server editing mode. */
export function NativePptxTableText({bytes,sourceRevision}:{bytes:Uint8Array;sourceRevision:string}){
 const active=useRef<AbortController|null>(null)
 const [result,setResult]=useState<{bytes:Uint8Array;inspection:NativePptxTableInspection}|null>(null)
 const [busy,setBusy]=useState(false),[error,setError]=useState('')
 useEffect(()=>{
  active.current?.abort();active.current=null;setResult(null);setBusy(false);setError('')
  return()=>{active.current?.abort();active.current=null}
 },[bytes,sourceRevision])
 async function inspect(){
  active.current?.abort()
  const controller=new AbortController();active.current=controller
  setBusy(true);setError('');setResult(null)
  let client:ReturnType<typeof createPptxWasmClient>|undefined
  try{
   client=createPptxWasmClient()
   const inspection=await client.inspectTables(bytes,{signal:controller.signal})
   if(inspection.source_revision!==sourceRevision)throw new Error('The table inspection does not match the opened presentation. Reopen the source and try again.')
   if(active.current===controller&&!controller.signal.aborted)setResult({bytes,inspection})
  }catch(reason){
   if(active.current===controller&&!controller.signal.aborted)setError(reason instanceof Error?reason.message:'Table source inspection was unavailable.')
  }finally{
   client?.terminate()
   if(active.current===controller){active.current=null;setBusy(false)}
  }
 }
 const current=result?.bytes===bytes&&result.inspection.source_revision===sourceRevision?result.inspection:null
 return <section aria-label="Browser-local presentation table inspection" className="ds-panel pptx-table-text">
  <h3>Table text</h3>
  <p>Read supported table cells without reconstructing their styling. This optional read-only view runs in your browser; no presentation bytes are uploaded.</p>
  <div className="ds-workstrip"><DsButton disabled={busy} onClick={()=>void inspect()}>Show read-only table text</DsButton>
   {busy&&<DsButton onClick={()=>{active.current?.abort();active.current=null;setBusy(false);setError('Inspection canceled. The original presentation is unchanged.')}}>Cancel table inspection</DsButton>}
  </div>
  {busy&&<p role="status">Reading table source in the browser…</p>}
  {error&&<p role="status">{error}</p>}
  {current&&<NativePptxTableTextResult inspection={current}/>}
 </section>
}
