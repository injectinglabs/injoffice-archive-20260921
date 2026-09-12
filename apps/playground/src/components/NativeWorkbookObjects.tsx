import {useEffect,useRef,useState} from 'react'
import {layoutNativeCachedChartV1,type NativeWorkbookObjectsV1} from '@injoffice/sheets/browser'
import {DsButton} from '../design-system/primitives'

export function NativeWorkbookObjects({bytes,revision,inspect,mode,onInspection}:{bytes:Uint8Array;revision:string;inspect:(bytes:Uint8Array,revision:string)=>Promise<NativeWorkbookObjectsV1>;mode:string;onInspection?:(result:NativeWorkbookObjectsV1|null)=>void}){
 const [result,setResult]=useState<NativeWorkbookObjectsV1|null>(null),[message,setMessage]=useState(''),[busy,setBusy]=useState(false)
 const generation=useRef(0)
 useEffect(()=>{generation.current++;setResult(null);setMessage('');setBusy(false);return()=>{generation.current++}},[bytes,revision,mode])
 async function run(){const token=++generation.current;setBusy(true);setResult(null);onInspection?.(null);setMessage('Reading saved chart data and table metadata…');try{const output=await inspect(bytes,revision);if(token!==generation.current)return;if(output.package_sha256!==revision)throw new Error('Object preview does not match this workbook');setResult(output);onInspection?.(output);setMessage(`${output.charts.length} chart parts and ${output.tables.length} table parts inspected. Read-only data preview; Office rendering fidelity is not established.`)}catch(e){if(token===generation.current)setMessage(e instanceof Error?e.message:'Object preview unavailable')}finally{if(token===generation.current)setBusy(false)}}
 return <section className="ds-panel" aria-label="Workbook charts and tables">
  <h3>Saved charts and table styles</h3>
  <p className="ds-muted">{mode==='server'?'This uploads the workbook to the configured helper when you choose the button.':'Inspection runs in the browser-local engine.'} Saved caches may be stale. No recalculation or external-link requests.</p>
  <DsButton disabled={busy} onClick={()=>void run()}>{busy?'Inspecting…':'Preview charts and table styles'}</DsButton><p role="status">{message}</p>
  {(result?.charts.length??0)>8&&<p role="note">Only the first 8 chart parts are plotted to bound browser rendering. Other chart parts: {result!.charts.slice(8).map(c=>c.part).join(', ')}.</p>}
  {result?.charts.slice(0,8).map((chart,index)=>{const layout=layoutNativeCachedChartV1(chart);return <section key={chart.part} aria-label={`Saved chart ${index+1}`}>
   <h4>Chart {index+1} · {layout?'cached data preview':'preview unavailable'}</h4><p className="ds-muted">{chart.part} · preview colors, not Office theme or positioning</p>
   {layout&&<svg role="img" aria-label={`Saved chart ${index+1}; values in the table below`} viewBox="0 0 600 260" style={{display:'block',width:'100%',maxWidth:700,border:'1px solid var(--rule)'}}><g transform="translate(45 15)">
    {chart.type==='col'?<line x1={0} x2={520} y1={layout.baseline*205} y2={layout.baseline*205} stroke="currentColor"/>:<line y1={0} y2={205} x1={layout.baseline*520} x2={layout.baseline*520} stroke="currentColor"/>}
    {layout.marks.map((mark,i)=><rect key={i} x={mark.x*520} y={mark.y*205} width={mark.width*520} height={mark.height*205} fill={['#3366CC','#CC6633','#339966','#993399'][mark.series%4]}><title>{chart.series[mark.series]?.name||`Series ${mark.series+1}`} · cache point {mark.point+1}: {mark.value}</title></rect>)}
   </g><text x={10} y={248} fontSize={12} fill="currentColor">Saved value range: {layout.minimum} to {layout.maximum}</text></svg>}
   {!layout&&chart.type!=='unsupported'&&<p>Plot unavailable: no usable values, unbounded numeric range, or more than 1,024 cache points. No subset of bars is substituted.</p>}
   {layout&&<details><summary>Saved chart values and series (first 100 of {chart.series.reduce((sum,s)=>sum+s.values.length,0)} values)</summary><table className="ds-table"><thead><tr><th>Series</th><th>Cache point</th><th>Saved value</th></tr></thead><tbody>{chart.series.flatMap((s,si)=>s.values.map((v,pi)=>({s,si,v,pi}))).slice(0,100).map(({s,si,v,pi})=><tr key={`${si}:${pi}`}><td>{s.name||`Unnamed series ${si+1}`}</td><td>{pi+1}</td><td>{v??'Missing'}</td></tr>)}</tbody></table></details>}
   <ul>{chart.warnings.map((warning,i)=><li key={i}>{warning}</li>)}</ul>
  </section>})}
  {!!result?.tables.length&&<details open><summary>Table styles — rendering still incomplete</summary><ul>{result.tables.map(table=><li key={table.part}>{table.name||table.part} · {table.ref} · {table.style||'No named style'} · {table.fill_preview?'Qualified fills only':'Style preview unavailable'}<ul>{table.warnings.map((warning,i)=><li key={i}>{warning}</li>)}</ul></li>)}</ul></details>}
 </section>
}
