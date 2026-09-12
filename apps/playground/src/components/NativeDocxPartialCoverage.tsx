import type {NativeDocxPartialContentV1} from '@injoffice/docs/native-docx'

/** Inventory-only UI: never creates editing targets or replaces richer view. */
export function NativeDocxPartialCoverage({coverage}:{coverage:NativeDocxPartialContentV1}){
 const labels={ 'unsupported-source':'Unsupported source content','unqualified-text-visibility':'Text visibility not qualified','hidden-text':'Hidden text','drawing':'Drawing not reconstructed','field':'Field result not reconstructed','reference':'Reference not reconstructed','control':'Document control not reconstructed','table':'Table layout not reconstructed','nonbody-story':'Header, footer or note not reconstructed','block-limit':'Body block limit','text-limit':'Text size limit','merged-cell':'Merged cell not reconstructed','cell-limit':'Table cell limit','nested-table':'Nested table content omitted','nested-table-limit':'Nested table omission limit'}
 return <section aria-label="Read-only projection coverage">
  <p className="native-muted ds-muted">Preview coverage: {coverage.coverage.visited_body_blocks} of {coverage.coverage.body_blocks} body blocks inspected; {coverage.source_diagnostics.document.length} source diagnostics retained. This separate read-only inventory does not measure Word pagination or change the richer editor below.</p>
  <details><summary>Read-only projection limitations ({coverage.omissions.length})</summary>
   <p>The reusable partial projection does not reproduce tables, drawings, fields or nonbody stories. Without a resolved style model it does not qualify text visibility. These are limits of that projection, not a claim that the editor below omits the same content.</p>
   <ul>{coverage.omissions.slice(0,20).map((item,index)=><li key={index}>{labels[item.code]} — {item.source.scope_id}{item.count>1?` (${item.count} units)`:''}</li>)}</ul>
   {coverage.blocks.filter(block=>block.kind==='table-source').map(table=><section key={table.source.scope_id}><h5>Table source coverage · not layout</h5><p>{table.cells.length} of {table.source_cell_count} cells inspected.</p><ul>{table.cells.slice(0,10).map(cell=><li key={cell.source.scope_id}>Source row {cell.row_ordinal+1}, cell {cell.cell_ordinal+1}: {cell.paragraphs.length} paragraph or omission groups</li>)}</ul>{table.cells.length>10&&<p>{table.cells.length-10} additional cell groups remain in the library inventory.</p>}</section>)}
   {coverage.omissions.length>20&&<p>{coverage.omissions.length-20} additional source-bound limitations remain in the library inventory.</p>}
  </details>
 </section>
}
