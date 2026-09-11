import {useEffect,useId,useRef,useState,type ReactNode} from 'react'
import {decodePptxPreview,type PptxPreview,type PreviewNode,type PreviewStroke} from '../../../pptx-page-paint-worker/src/contract'
import {readNativePreviewResponse,decodeNativeDocxImages,nativeDocxImagesWithinBudget} from './NativeDocxPages'
import {DsButton,DsField,DsSelect} from '../design-system/primitives'
import type {NativePptxDeck} from '@injoffice/pptx-native'
import {nativePptxPreviewStatus} from '../nativePptxPreviewStatus'

export function NativePptxVector({preview,onImageError}:{preview:PptxPreview;onImageError?:()=>void}){
 const prefix=useId().replace(/:/g,'')
 const color=(value:string)=>value==='none'?'none':`#${value}`
 const strokeProps=(node:PreviewStroke)=>({strokeLinecap:node.strokeLinecap,strokeLinejoin:node.strokeLinejoin,strokeMiterlimit:node.strokeMiterlimit})
 function draw(node:PreviewNode,key:string):ReactNode{
  switch(node.kind){
   case 'image':{const resource=preview.resources.find(r=>r.id===node.resourceId)!;const c=node.crop??{left:0,top:0,right:0,bottom:0};return <svg key={key} x={node.rect.x} y={node.rect.y} width={node.rect.cx} height={node.rect.cy} viewBox={`${c.left} ${c.top} ${100000-c.left-c.right} ${100000-c.top-c.bottom}`} preserveAspectRatio="none" overflow="hidden"><image data-native-raster={resource.id} href={`data:${resource.content_type};base64,${resource.bytes_base64}`} width={100000} height={100000} preserveAspectRatio="none" onError={onImageError}/></svg>}
   case 'group':{const clip=node.clip,id=`${prefix}-${key}`;return <g key={key} data-native-source-role={node.sourceRole} transform={`matrix(${node.transform.join(' ')})`}>
    {clip&&<defs><clipPath id={id}><rect x={clip.x} y={clip.y} width={clip.cx} height={clip.cy}/></clipPath></defs>}
    <g clipPath={clip?`url(#${id})`:undefined}>{node.children.map((child,i)=>draw(child,`${key}-${i}`))}</g>
   </g>}
   case 'path':return <path key={key} d={node.d} fill={color(node.fill)} stroke={node.stroke?color(node.stroke):undefined} strokeWidth={node.strokeWidth} {...strokeProps(node)}/>
   case 'rect':return <rect key={key} x={node.rect.x} y={node.rect.y} width={node.rect.cx} height={node.rect.cy} rx={node.radius} fill={color(node.fill)} stroke={node.stroke?color(node.stroke):undefined} strokeWidth={node.strokeWidth} {...strokeProps(node)}/>
   case 'ellipse':return <ellipse key={key} cx={node.rect.x+node.rect.cx/2} cy={node.rect.y+node.rect.cy/2} rx={node.rect.cx/2} ry={node.rect.cy/2} fill={color(node.fill)} stroke={node.stroke?color(node.stroke):undefined} strokeWidth={node.strokeWidth} {...strokeProps(node)}/>
   case 'placeholder':return <g key={key}><rect x={node.rect.x} y={node.rect.y} width={node.rect.cx} height={node.rect.cy} fill="#eee" stroke="#777" strokeWidth={12700}/><text x={node.rect.x+12700} y={node.rect.y+127000} fontSize={101600}>{node.label}</text></g>
  }
 }
 return <svg role="img" aria-label={`Measured native slide ${preview.slide_index+1}`} viewBox={`0 0 ${preview.width} ${preview.height}`} style={{display:'block',width:'100%',background:color(preview.background),border:'1px solid var(--ds-line)'}}>{preview.nodes.map((node,i)=>draw(node,String(i)))}</svg>
}

export function NativePptxPreviewResult({preview,source,onImageError}:{preview:PptxPreview;source?:NativePptxDeck;onImageError?:()=>void}){
 const coverage=nativePptxPreviewStatus(preview,source)
 const limit=50
 return <section aria-label={`Native preview result for slide ${preview.slide_index+1}`} data-native-preview-status={coverage.status}>
  <h4>{coverage.label} · slide {preview.slide_index+1} of {preview.slide_count}</h4>
  <p className="ds-muted">Supported paint is retained alongside identified gaps. Native line layout uses the InjOffice policy; complete source coverage and PowerPoint equivalence are not established. Read-only; the original file and editing permissions are unchanged.</p>
  <NativePptxVector preview={preview} onImageError={onImageError}/>
  {coverage.reasons.length>0&&<details open><summary>Missing content and coverage limits ({coverage.reasons.length})</summary><ul>{coverage.reasons.slice(0,limit).map((reason,i)=><li key={i}>{reason}</li>)}</ul>{coverage.reasons.length>limit&&<p>{coverage.reasons.length-limit} additional coverage reasons are not displayed.</p>}</details>}
  <details><summary>Native paint counts</summary><p>{coverage.paintPrimitives} paint records · {coverage.glyphRecords} text glyph records · {coverage.placeholderRegions} placeholder regions{coverage.sourceBound?` · ${coverage.knownRefusedObjects} known refused source objects`:''}. Paint records are not object counts or proof of visible pixels after clipping.</p></details>
  {preview.diagnostics.length>0&&<details><summary>Native diagnostics ({preview.diagnostics.length})</summary><ul>{preview.diagnostics.slice(0,limit).map((message,i)=><li key={i}>{message}</li>)}</ul>{preview.diagnostics.length>limit&&<p>{preview.diagnostics.length-limit} additional diagnostics are not displayed.</p>}</details>}
 </section>
}

export function NativePptxSlides({bytes,slideCount,apiBase,source}:{bytes:Uint8Array;slideCount:number;apiBase:string;source?:NativePptxDeck}){
 const [paint,setPaint]=useState<PptxPreview|null>(null),[busy,setBusy]=useState(false),[at,setAt]=useState(0)
 const consent=`Native slides require uploading this presentation to ${apiBase}. Nothing is uploaded until you choose the button below.`
 const [message,setMessage]=useState(consent)
 const generation=useRef(0),pending=useRef<AbortController|null>(null)
 useEffect(()=>{generation.current++;pending.current?.abort();setPaint(null);setBusy(false);setAt(0);setMessage(consent);return()=>{generation.current++;pending.current?.abort()}},[bytes,apiBase])
 async function render(){
  pending.current?.abort()
  const token=++generation.current,controller=new AbortController();pending.current=controller
  setBusy(true);setPaint(null);setMessage('Shaping this slide with exact operator-provided font bytes…')
  try{
   const owned=Uint8Array.from(bytes),hash=await crypto.subtle.digest('SHA-256',owned.buffer),digest=[...new Uint8Array(hash)].map(n=>n.toString(16).padStart(2,'0')).join('')
   if(controller.signal.aborted||token!==generation.current)return
   const response=await fetch(`${apiBase}/v1/pptx/slide-preview?slide=${at}`,{method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.presentationml.presentation'},body:new Blob([owned.buffer]),credentials:'omit',redirect:'error',signal:controller.signal})
   const result=await readNativePreviewResponse(response)
   if(!response.ok)throw new Error(typeof (result as {error?:unknown})?.error==='string'?(result as {error:string}).error:'Native slide preview was refused by the helper.')
   const decoded=decodePptxPreview(result)
   if(decoded.package_sha256!==digest||decoded.slide_index!==at||decoded.slide_count!==slideCount)throw new Error('Native slide does not match the opened source.')
   if(!nativeDocxImagesWithinBudget(decoded.resources))throw new Error('Native slide image pixel budget exceeded.')
   await decodeNativeDocxImages(decoded.resources,controller.signal)
   if(controller.signal.aborted||token!==generation.current)return
   setPaint(decoded);setMessage(`${nativePptxPreviewStatus(decoded,source).label}. Review the content and coverage limits below. The source file is unchanged.`)
  }catch(error){if(!controller.signal.aborted&&token===generation.current)setMessage(`${error instanceof Error?error.message:'Native preview failed'} The file preview remains available; the original file is unchanged.`)}
  finally{if(token===generation.current)setBusy(false)}
 }
 function changeSlide(index:number){generation.current++;pending.current?.abort();setBusy(false);setPaint(null);setAt(index);setMessage(consent)}
 return <section aria-label="Measured native presentation" className="ds-panel">
  <h3>Measured native slide</h3><p className="ds-status" role="status">{message}</p>
  <div className="ds-workstrip"><DsField label="Native slide"><DsSelect value={at} onChange={event=>changeSlide(Number(event.target.value))}>{Array.from({length:Math.min(slideCount,10000)},(_,i)=><option key={i} value={i}>{i+1} of {slideCount}</option>)}</DsSelect></DsField>
  <DsButton disabled={busy} onClick={()=>void render()}>{busy?'Rendering native slide…':'Upload to helper and render native slide'}</DsButton></div>
  {paint&&<NativePptxPreviewResult preview={paint} source={source} onImageError={()=>{setPaint(null);setMessage('Native image decoding failed. Native rendering was cleared; the original source is unchanged.')}}/>}
 </section>
}
