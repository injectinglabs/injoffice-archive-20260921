import {createElement,useEffect,useRef,useState} from 'react'
import {createDocxWasmClient} from '@injoffice/docx-wasm'
import {createNativeDocxPartialContentPreviewV1,type NativeDocxPartialContentV1,type NativeDocxPartialParagraphV1,type NativeDocxMathNodeV1,type NativeDocxEquationPreviewV1} from '@injoffice/docs/native-docx'
import {DsButton} from '../design-system/primitives'

function Paragraph({paragraph}:{paragraph:NativeDocxPartialParagraphV1}){
 return <p style={{whiteSpace:'pre-wrap'}}>{paragraph.segments.map((segment,index)=>segment.kind==='omission'?<span key={index}>[{segment.code}]</span>:segment.kind==='alternative-text'?<span key={index}>[Authored drawing description: {segment.text}]</span>:<span key={index}>{segment.text}</span>)}</p>
}
function MathNode({node}:{node:NativeDocxMathNodeV1}):React.ReactNode{
 if(node.kind==='text')return createElement('mtext',null,node.text)
 const tag={row:'mrow',fraction:'mfrac',superscript:'msup',subscript:'msub',radical:'msqrt'}[node.kind]
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
  {preview.retained_nontext_diagnostic_ids.length>0&&<p>{preview.retained_nontext_diagnostic_ids.length} source metadata warnings are retained. They do not prevent plain text recovery, but this view does not reproduce their formatting.</p>}
  {preview.blocks.map((block,index)=>block.kind==='paragraph'?<Paragraph key={index} paragraph={block}/>:block.kind==='omission'?<p key={index}>[{block.code}: {block.count}]</p>:<section key={index} aria-label="Table source text"><h5>Table cell text · no table layout</h5>{block.cells.map(cell=><section key={cell.source.scope_id}><h6>Source row {cell.row_ordinal+1}, cell {cell.cell_ordinal+1}</h6>{cell.source_merge&&<p>Authored span: {cell.source_merge.grid_span} grid {cell.source_merge.grid_span===1?'column':'columns'}. {cell.source_merge.vertical_merge==='continue'?'Vertical merge continuation — text omitted.':cell.source_merge.vertical_merge==='restart'?'Vertical merge starts here — owner text only.':'Horizontal merge — owner text only.'} Merge geometry is not reproduced.</p>}{cell.paragraphs.map((paragraph,n)=>paragraph.kind==='paragraph'?<Paragraph key={n} paragraph={paragraph}/>:<p key={n}>[{paragraph.code}]</p>)}</section>)}</section>)}
 </article>
}

/** Separate opt-in browser operation, including while editing in server mode. */
export function NativeDocxPartialText({bytes,packageDigest}:{bytes:Uint8Array;packageDigest:string}){
 const active=useRef<AbortController|null>(null)
 const [result,setResult]=useState<NativeDocxPartialContentV1|null>(null),[busy,setBusy]=useState(false),[error,setError]=useState('')
 const [equations,setEquations]=useState<NativeDocxEquationPreviewV1[]>([])
 useEffect(()=>{active.current?.abort();active.current=null;setResult(null);setBusy(false);setError('');return()=>{active.current?.abort();active.current=null}},[bytes,packageDigest])
 const run=async()=>{
  const controller=new AbortController();active.current?.abort();active.current=controller
  setBusy(true);setError('');setResult(null);setEquations([])
  let client:ReturnType<typeof createDocxWasmClient>|undefined
  try{
   client=createDocxWasmClient()
   const joined=await client.inspectPartialContent(bytes,{signal:controller.signal})
   if(joined.document.source.package_sha256!==packageDigest)throw new Error('Source changed; reopen the partial text preview.')
   const preview=createNativeDocxPartialContentPreviewV1(joined.document,{policy:'source-text-with-omissions-v1',read_only:true},joined.resolved_layout,joined.nested_table_omissions)
   if(active.current===controller&&!controller.signal.aborted){setResult(preview);setEquations(joined.equations??[])}
  }catch(reason){if(active.current===controller&&!controller.signal.aborted)setError(reason instanceof Error?reason.message:'Partial source text could not be qualified.')}
  finally{client?.terminate();if(active.current===controller){active.current=null;setBusy(false)}}
 }
 return <section aria-label="Browser-local read-only partial text">
  <p>This optional text-only view resolves source styles in your browser. No file is uploaded, including when the editor uses server mode.</p>
  <DsButton disabled={busy} onClick={()=>void run()}>Show read-only partial text</DsButton>
  {busy&&<DsButton onClick={()=>{active.current?.abort();active.current=null;setBusy(false)}}>Cancel partial text</DsButton>}
  {busy&&<p role="status">Reading source text in the browser…</p>}
  {error&&<p role="status">{error}</p>}
  {result?.source.package_sha256===packageDigest&&<><NativeDocxPartialTextView preview={result}/><NativeDocxEquationList equations={equations}/></>}
 </section>
}
