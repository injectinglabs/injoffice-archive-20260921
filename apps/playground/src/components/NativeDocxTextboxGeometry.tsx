import {useEffect,useRef,useState} from 'react'
import {createDocxWasmClient} from '@injoffice/docx-wasm'
import {decodeNativeDocxTextboxShapePaintV1,nativeTextboxFontDigestV1,type NativeDocxTextboxShapePaintV1} from '@injoffice/docs/native-docx'

type Joined=Awaited<ReturnType<ReturnType<typeof createDocxWasmClient>['inspectPartialContent']>>
export function NativeDocxTextboxShapeView({paint}:{paint:NativeDocxTextboxShapePaintV1}){
 if(paint.status==='omitted')return <p>[Rectangle preview omitted: {paint.reason}]</p>
 const w=paint.width_millipoints,h=paint.height_millipoints,pad=paint.line_width_millipoints/2
 return <figure><svg aria-label="Authored rectangle textbox preview" role="img" viewBox={`${-pad} ${-pad} ${w+2*pad} ${h+2*pad}`} width={`${(w+2*pad)/1000}pt`} height={`${(h+2*pad)/1000}pt`} style={{maxWidth:'100%',overflow:'visible'}}><rect x={0} y={0} width={w} height={h} fill={paint.fill_rgb==='none'?'none':'#'+paint.fill_rgb} stroke={paint.line_rgb==='none'?'none':'#'+paint.line_rgb} strokeWidth={paint.line_width_millipoints} strokeLinejoin="miter" strokeMiterlimit={8}/>{paint.paths.map((d,i)=><path key={i} d={d} fill={'#'+paint.text_rgb}/>)}</svg><figcaption>Authored rectangle: {w/1000} × {h/1000} pt. Text uses the supplied font’s glyph outlines. {paint.line_layout&&<>{paint.line_layout.lines.length} authored lines. Each authored line is centered within its specified line height. </>}{paint.wrap_paint&&<>{paint.wrap_paint.lines.length} lines wrap at spaces within the specified width. Spaces remain in the measured line width. </>}Page placement is not produced.</figcaption></figure>
}
export function NativeDocxTextboxGeometry({bytes,packageDigest}:{bytes:Uint8Array;packageDigest:string}){
 const [joined,setJoined]=useState<Joined|null>(null),[font,setFont]=useState<Uint8Array|null>(null),[consent,setConsent]=useState(false),[busy,setBusy]=useState(false),[error,setError]=useState(''),[paint,setPaint]=useState<NativeDocxTextboxShapePaintV1|null>(null),[index,setIndex]=useState(0)
 const active=useRef<AbortController|null>(null),fontEpoch=useRef(0)
 useEffect(()=>{active.current?.abort();fontEpoch.current++;setFont(null);setJoined(null);setPaint(null);setConsent(false);setBusy(false);setError('');return()=>active.current?.abort()},[bytes,packageDigest])
 const inspect=async()=>{
  const c=new AbortController();active.current?.abort();active.current=c;setBusy(true);setError('');setPaint(null);setJoined(null)
  const client=createDocxWasmClient()
  try{const value=await client.inspectPartialContent(bytes,{signal:c.signal});if(value.document.source.package_sha256!==packageDigest)throw new Error('Source changed');if(!c.signal.aborted){setJoined(value);setIndex(0)}}catch(e){if(!c.signal.aborted)setError(e instanceof Error?e.message:'Inspection failed')}finally{client.terminate();if(!c.signal.aborted)setBusy(false)}
 }
 const render=async()=>{
  if(!joined||!font||!consent)return
  const c=new AbortController();active.current?.abort();active.current=c;setBusy(true);setError('');setPaint(null)
  try{let binary='';for(let i=0;i<font.length;i+=8192)binary+=String.fromCharCode(...font.subarray(i,i+8192))
   const response=await fetch('/__injoffice/textbox-geometry',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({document:joined.document,evidence:joined.textbox_geometry,index,font_base64:btoa(binary)}),signal:c.signal})
   if(!response.ok)throw new Error('The Node preview helper is unavailable or refused this request. Use the local development helper or the package compiler API.')
   const text=await response.text();if(text.length>12000000)throw new Error('Preview response exceeded its budget')
   const result=decodeNativeDocxTextboxShapePaintV1(joined.document,joined.textbox_geometry,index,JSON.parse(text),nativeTextboxFontDigestV1(font));if(!c.signal.aborted)setPaint(result)
  }catch(e){if(!c.signal.aborted)setError(e instanceof Error?e.message:'Preview failed')}finally{if(!c.signal.aborted)setBusy(false)}
 }
 return <section aria-label="Read-only rectangle textbox geometry">
  <h4>Rectangle textbox preview</h4><p>Inspect explicit DrawingML rectangles, then preview supplied-font text inside the authored bounds, including up to 16 explicit authored lines or automatically wrapped ASCII word lines. No page placement, autofit or editing is provided.</p>
  <button type="button" disabled={busy} onClick={()=>void inspect()}>Inspect rectangle geometry in browser</button>
  {joined&&<><p>{joined.textbox_geometry?.items.length??0} rectangle candidates. Original drawing warnings remain.</p>{!!joined.textbox_geometry?.items.length&&<>
   <label>Textbox <select disabled={busy} value={index} onChange={e=>{setIndex(Number(e.target.value));setPaint(null)}}>{joined.textbox_geometry.items.map((item,i)=><option key={item.owner.diagnostic_id} value={i}>{i+1}: {item.geometry?.font_family??item.owner.reason}</option>)}</select></label>
   <label>Supply the exact regular TrueType font <input type="file" disabled={busy} accept=".ttf" onChange={e=>{const file=e.target.files?.[0],epoch=++fontEpoch.current;setPaint(null);setFont(null);if(file&&file.size<=16*1024*1024)void file.arrayBuffer().then(b=>{if(epoch===fontEpoch.current)setFont(new Uint8Array(b))}).catch(()=>{if(epoch===fontEpoch.current)setError("Font could not be read")});else setError('Supply a font no larger than 16 MiB.')}}/></label>
   <p>The Node helper receives the extracted document model (including body, headers, notes and comments), textbox text and geometry, and the supplied font bytes. The original DOCX bytes stay in this browser. This demo uses the same-origin development helper; it is not a browser-only rendering operation.</p>
   <label><input type="checkbox" disabled={busy} checked={consent} onChange={e=>setConsent(e.target.checked)}/>Allow sending that content and font to the preview helper</label>
   <button type="button" disabled={busy||!font||!consent} onClick={()=>void render()}>Render authored rectangle</button>
  </>}</>}
  {busy&&<p role="status">Qualifying rectangle preview…</p>}{error&&<p role="status">{error}</p>}{paint?.package_sha256===packageDigest&&<NativeDocxTextboxShapeView paint={paint}/>}
 </section>
}
