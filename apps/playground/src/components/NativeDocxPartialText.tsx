import {createNativeDocxTextboxInventoryV1,type NativeDocxTextboxInventoryV1} from '@injoffice/docs/native-docx'
import {createElement,useEffect,useRef,useState} from 'react'
import {createDocxWasmClient} from '@injoffice/docx-wasm'
import {createNativeDocxReviewInventoryV1,type NativeDocxReviewInventoryV1,createNativeDocxPartialContentPreviewV1,type NativeDocxPartialContentV1,type NativeDocxPartialParagraphV1,type NativeDocxMathNodeV1,type NativeDocxEquationPreviewV1} from '@injoffice/docs/native-docx'
import {DsButton} from '../design-system/primitives'

function Paragraph({paragraph}:{paragraph:NativeDocxPartialParagraphV1}){
 return <p style={{whiteSpace:'pre-wrap'}}>{paragraph.segments.map((segment,index)=>segment.kind==='omission'?<span key={index}>[{segment.code}]</span>:segment.kind==='alternative-text'?<span key={index}>[Authored drawing description: {segment.text}]</span>:<span key={index}>{segment.text}</span>)}</p>
}
function MathNode({node}:{node:NativeDocxMathNodeV1}):React.ReactNode{
 if(node.kind==='text')return createElement('mtext',null,node.text)
 const tag={row:'mrow',fraction:'mfrac',superscript:'msup',subscript:'msub',radical:'msqrt','indexed-radical':'mroot'}[node.kind]
 return createElement(tag,null,node.children.map((child,index)=><MathNode key={index} node={child}/>))
}
export function NativeDocxEquationList({equations}:{equations:NativeDocxEquationPreviewV1[]}){
 if(!equations.length)return null
 const noticeCount=new Set(equations.flatMap(e=>e.context_notice_ids??[])).size
 return <section aria-label="Read-only equation approximations"><h4>Equation previews</h4><p>Browser math layout, not Word typography or pagination. Equations are listed separately in source order; the original unsupported-source warnings remain.</p>{noticeCount>0&&<p>{noticeCount} source context notices retained: this equation-only view does not reproduce paragraph tabs, hyphenation, section layout or source font matching. Already decoded Unicode is unchanged.</p>}{equations.map((equation,index)=><div key={equation.diagnostic_id}><p>Equation {index+1}</p>{equation.status==='supported'&&equation.tree?createElement('math',{xmlns:'http://www.w3.org/1998/Math/MathML',display:'block'},<MathNode node={equation.tree}/>):<p>[Equation omitted: {equation.reason}]</p>}</div>)}</section>
}
export function NativeDocxPartialTextView({preview}:{preview:NativeDocxPartialContentV1}){
 return <article aria-label="Read-only partial source text">
  <p>Source text only, not Word layout. Table cells are listed in source order; omissions remain labeled. This view cannot edit the document.</p>
  {!!preview.table_text_contexts?.length&&<p>{preview.table_text_contexts.length} source-qualified geometry-only table style contexts allow plain text recovery. Conditional styles remain unsupported; table look, borders, spacing and margins are not reconstructed.</p>}
  {preview.retained_nontext_diagnostic_ids.length>0&&<p>{preview.retained_nontext_diagnostic_ids.length} source metadata warnings are retained. They do not prevent plain text recovery, but this view does not reproduce their formatting.</p>}
  {preview.blocks.map((block,index)=>block.kind==='paragraph'?<Paragraph key={index} paragraph={block}/>:block.kind==='omission'?<p key={index}>[{block.code}: {block.count}]</p>:<section key={index} aria-label="Table source text"><h5>Table cell text · no table layout</h5>{block.cells.map(cell=><section key={cell.source.scope_id}><h6>Source row {cell.row_ordinal+1}, cell {cell.cell_ordinal+1}</h6>{cell.source_merge&&<p>Authored span: {cell.source_merge.grid_span} grid {cell.source_merge.grid_span===1?'column':'columns'}. {cell.source_merge.vertical_merge==='continue'?'Vertical merge continuation — text omitted.':cell.source_merge.vertical_merge==='restart'?'Vertical merge starts here — owner text only.':'Horizontal merge — owner text only.'} Merge geometry is not reproduced.</p>}{cell.paragraphs.map((paragraph,n)=>paragraph.kind==='paragraph'?<Paragraph key={n} paragraph={paragraph}/>:<p key={n}>[{paragraph.code}]</p>)}</section>)}</section>)}
  {!!preview.header_footer_stories?.length&&<section aria-label="Header and footer source inventory"><h4>Header and footer source text</h4><p>Source inventory only. No active first, even or default variant is selected, and no page placement or repetition is reproduced. Tables remain omitted.</p>{preview.header_footer_stories.map(story=><section key={story.source.scope_id}><h5>{story.kind==='header'?'Header':'Footer'} source · {story.source.anchor.part_name}</h5>{story.blocks.map((block,index)=>block.kind==='paragraph'?<Paragraph key={index} paragraph={block}/>:<p key={index}>[{block.code}: {block.count}]</p>)}</section>)}</section>}
  {preview.comment_inventory&&<section aria-label="Read-only comment source inventory"><h4>Comment source text</h4><p>Authored comments are listed separately in source order. Authors and dates are stored metadata. Comment ranges, threads, revision display and tables are not reconstructed. This view cannot accept or reject changes.</p>{preview.comment_inventory.stories.length===0&&<p>No qualified comment stories are available; source omissions remain listed.</p>}{preview.comment_inventory.stories.map(story=><section key={story.source.scope_id}><h5>Comment {story.native_comment_id} · {story.author}</h5>{story.created_at&&<p>Stored date: {story.created_at}</p>}{story.blocks.map((block,index)=>block.kind==='paragraph'?<Paragraph key={index} paragraph={block}/>:<p key={index}>[{block.code}: {block.count}]</p>)}</section>)}</section>}
 </article>
}

export function NativeDocxReviewInventoryView({review}:{review:NativeDocxReviewInventoryV1}){
 const diagnosticCodes=[...new Set([...review.source_diagnostics.document,...review.source_diagnostics.resolved].map(d=>d.code))]
 return <section aria-label="Read-only tracked-change source inventory"><h4>Tracked-change source inventory</h4><p>Stored revision metadata, not verified author identities. Changes are listed separately; this is not a final or original Word view. Only qualified insertion text is shown. Deleted and moved text stays omitted. No accept/reject or editing controls are provided.</p>{diagnosticCodes.length>0&&<details><summary>Retained source diagnostic codes ({diagnosticCodes.length})</summary><ul>{diagnosticCodes.map(code=><li key={code}>{code}</li>)}</ul></details>}{review.items.length===0&&<p>No qualified review wrappers are available.</p>}{review.items.map(item=><section key={item.diagnostic_id}><h5>{item.kind} {item.revision_id} · {item.author}</h5>{item.created_at&&<p>Stored date: {item.created_at}</p>}{item.text_status==='qualified-insertion'?<p style={{whiteSpace:'pre-wrap'}}>{item.segments.map((s,i)=>s.kind==='text'?<span key={i}>{s.text}</span>:null)}</p>:<p>[Review text omitted: source structure or visibility is unqualified, or this change kind is metadata only.]</p>}<details><summary>Retained source diagnostics ({item.retained_diagnostic_ids.length})</summary><ul>{item.retained_diagnostic_ids.map(id=><li key={id}>{id}</li>)}</ul></details></section>)}{review.omitted_count>0&&<p>[Review inventory omissions: {review.omitted_count}]</p>}</section>
}

export function NativeDocxTextboxInventoryView({inventory}:{inventory:NativeDocxTextboxInventoryV1}){
 return <section aria-label="Read-only textbox source inventory"><h4>Textbox source text</h4>
  <p>Textboxes are listed separately in source order. Shape geometry, placement, wrapping and links between textboxes are not reconstructed. Source omissions remain labeled.</p>
  {inventory.items.length===0&&<p>No qualified textbox inventory is available; source drawing warnings remain.</p>}
  {inventory.items.map((item,index)=><section key={item.diagnostic_id}><h5>Textbox {index+1} · {item.kind==='vml'?'VML':'DrawingML'}</h5>{item.status==='supported'?item.paragraphs.map((text,n)=><p key={n} style={{whiteSpace:'pre-wrap'}}>{text}</p>):<p>[Textbox text omitted: {item.reason}]</p>}</section>)}
  {inventory.omitted_count>0&&<p>[Textbox inventory limit: {inventory.omitted_count} omitted]</p>}
  <details><summary>Retained source warnings ({inventory.source_diagnostics.length})</summary><ul>{inventory.source_diagnostics.map(d=><li key={d.id}>{d.code}: {d.message}</li>)}</ul></details>
 </section>
}

/** Separate opt-in browser operation, including while editing in server mode. */
export function NativeDocxPartialText({bytes,packageDigest}:{bytes:Uint8Array;packageDigest:string}){
 const active=useRef<AbortController|null>(null)
 const [result,setResult]=useState<NativeDocxPartialContentV1|null>(null),[busy,setBusy]=useState(false),[error,setError]=useState('')
 const [textboxes,setTextboxes]=useState<NativeDocxTextboxInventoryV1|null>(null)
 const [review,setReview]=useState<NativeDocxReviewInventoryV1|null>(null)
 const [equations,setEquations]=useState<NativeDocxEquationPreviewV1[]>([])
 useEffect(()=>{active.current?.abort();active.current=null;setResult(null);setTextboxes(null);setReview(null);setEquations([]);setBusy(false);setError('');return()=>{active.current?.abort();active.current=null}},[bytes,packageDigest])
 const run=async(includeComments=false,includeReview=false,includeTextboxes=false)=>{
  const controller=new AbortController();active.current?.abort();active.current=controller
  setBusy(true);setError('');setResult(null);setEquations([]);setReview(null);setTextboxes(null)
  let client:ReturnType<typeof createDocxWasmClient>|undefined
  try{
   client=createDocxWasmClient()
   const joined=await client.inspectPartialContent(bytes,{signal:controller.signal})
   if(joined.document.source.package_sha256!==packageDigest)throw new Error('Source changed; reopen the partial text preview.')
   const preview=createNativeDocxPartialContentPreviewV1(joined.document,{policy:'source-text-with-omissions-v1',read_only:true,...(includeComments?{comment_policy:'source-comment-inventory-v1' as const}:{})},joined.resolved_layout,joined.nested_table_omissions,joined.table_text_contexts)
   if(active.current===controller&&!controller.signal.aborted){setResult(preview);setTextboxes(includeTextboxes?createNativeDocxTextboxInventoryV1(joined.document,{policy:'source-textbox-inventory-v1',read_only:true},joined.textbox_inventory):null);setEquations(joined.equations??[]);setReview(includeReview?createNativeDocxReviewInventoryV1(joined.document,{policy:'source-review-inventory-v1',read_only:true},joined.resolved_layout,joined.review_changes):null)}
  }catch(reason){if(active.current===controller&&!controller.signal.aborted)setError(reason instanceof Error?reason.message:'Partial source text could not be qualified.')}
  finally{client?.terminate();if(active.current===controller){active.current=null;setBusy(false)}}
 }
 return <section aria-label="Browser-local read-only partial text">
  <p>This optional text-only view resolves source styles in your browser. No file is uploaded, including when the editor uses server mode.</p>
  <DsButton disabled={busy} onClick={()=>void run()}>Show read-only partial text</DsButton>
  <DsButton disabled={busy} onClick={()=>void run(true)}>Show partial text with comments</DsButton>
  <DsButton disabled={busy} onClick={()=>void run(false,true)}>Inspect tracked-change source</DsButton>
  <DsButton disabled={busy} onClick={()=>void run(false,false,true)}>Inspect textbox source</DsButton>
  {busy&&<DsButton onClick={()=>{active.current?.abort();active.current=null;setBusy(false)}}>Cancel partial text</DsButton>}
  {busy&&<p role="status">Reading source text in the browser…</p>}
  {error&&<p role="status">{error}</p>}
  {result?.source.package_sha256===packageDigest&&<><NativeDocxPartialTextView preview={result}/><NativeDocxEquationList equations={equations}/>{textboxes?.source.package_sha256===packageDigest&&<NativeDocxTextboxInventoryView inventory={textboxes}/>}{review?.source.package_sha256===packageDigest&&<NativeDocxReviewInventoryView review={review}/>}</>}
 </section>
}
