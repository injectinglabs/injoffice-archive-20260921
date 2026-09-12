import type {NativeDocxPartialContentV1} from '@injoffice/docs/native-docx'

/** Inventory-only UI: never creates editing targets or replaces richer view. */
export function NativeDocxPartialCoverage({coverage}:{coverage:NativeDocxPartialContentV1}){
 const labels={ 'unsupported-source':'Unsupported source content','unqualified-text-visibility':'Text visibility not qualified','hidden-text':'Hidden text','drawing':'Drawing not reconstructed','field':'Field result not reconstructed','reference':'Reference not reconstructed','control':'Document control not reconstructed','table':'Table layout not reconstructed','nonbody-story':'Header, footer or note not reconstructed','block-limit':'Body block limit','text-limit':'Text size limit'}
 return <section aria-label="Read-only projection coverage">
  <p className="native-muted ds-muted">Preview coverage: {coverage.coverage.visited_body_blocks} of {coverage.coverage.body_blocks} body blocks inspected; {coverage.source_diagnostics.document.length} source diagnostics retained. This separate read-only inventory does not measure Word pagination or change the richer editor below.</p>
  <details><summary>Read-only projection limitations ({coverage.omissions.length})</summary>
   <p>The reusable partial projection does not reproduce tables, drawings, fields or nonbody stories. Without a resolved style model it does not qualify text visibility. These are limits of that projection, not a claim that the editor below omits the same content.</p>
   <ul>{coverage.omissions.slice(0,20).map((item,index)=><li key={index}>{labels[item.code]} — {item.source.scope_id}{item.count>1?` (${item.count} units)`:''}</li>)}</ul>
   {coverage.omissions.length>20&&<p>{coverage.omissions.length-20} additional source-bound limitations remain in the library inventory.</p>}
  </details>
 </section>
}
