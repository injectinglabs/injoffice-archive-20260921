/** Browser-safe composition of every source-qualified page textbox. */
import {resolveTextboxPosition} from './nativeTextboxPositionV2.js'
import {decodeNativeDocxDocument} from './nativeContract.js'
import {decodeNativeDocxPagePaintV1} from './nativePagePaintWireV1.js'
import type {NativeDocxPagePaintSuccessV1} from './nativePagePaintV1.js'
import {decodeNativeDocxTextboxGeometryV1,decodeNativeDocxTextboxShapePaintV1,nativeTextboxGeometryDigestV1,nativeTextboxGeometryPlainData,type NativeDocxTextboxShapePaintV1} from './nativeTextboxGeometryPreviewV1.js'

export const DOCX_TEXTBOX_PAGES_PREVIEW_PROTOCOL='injoffice.docx.textbox-page-preview' as const
export const DOCX_TEXTBOX_PAGES_PREVIEW_WARNING='Read-only approximate textbox composition in source order; original drawing restrictions remain.' as const

export function projectNativeDocxTextboxPagesV2(source:unknown,evidence:unknown){
 const decoded=decodeNativeDocxDocument(source)
 if(!decoded.ok)throw new TypeError('Invalid textbox page source')
 const document=structuredClone(decoded.value),geometry=decodeNativeDocxTextboxGeometryV1(document,evidence)
 if(!geometry.items.length||geometry.items.length>64||geometry.omitted_count!==0)throw new TypeError('Textbox pages require complete rectangle evidence')
 const paragraphs=new Map(document.body.blocks.flatMap(b=>b.paragraph?[[b.id,b.paragraph] as const]:[]))
 const diagnostics=new Map(document.unsupported.map(d=>[d.id,d])),selected=new Set<string>()
 const items=geometry.items.map((item,index)=>{
  const position=item.page_anchor,p=paragraphs.get(item.owner.paragraph_id),diagnostic=diagnostics.get(item.owner.diagnostic_id)
  if(!position||!item.geometry||!p||!(diagnostic?.code==='PICTURE_GRAPHIC_REQUIRED'||(position.policy==='relative-position-no-wrap-v2'&&diagnostic?.code==='FLOATING_DRAWING_SEMANTICS_PRESERVED'))||selected.has(diagnostic.id))throw new TypeError('Unsupported textbox page source')
  const path=position.source_anchor.path.slice(p.anchor.path.length)
  if(!position.source_anchor.path.startsWith(p.anchor.path)||!/^\/w:r\[[1-9][0-9]*\]\/w:drawing\[[1-9][0-9]*\]\//u.test(path)||p.runs.some(r=>r.anchor.start_byte<=position.source_anchor.end_byte))throw new TypeError('Textboxes must precede all text in their anchor paragraph')
  selected.add(diagnostic.id)
  return {item,index}
 }).sort((a,b)=>a.item.page_anchor!.source_anchor.start_byte-b.item.page_anchor!.source_anchor.start_byte)
 for(let i=1;i<items.length;i++)if(items[i-1]!.item.page_anchor!.source_anchor.end_byte>=items[i]!.item.page_anchor!.source_anchor.start_byte)throw new TypeError('Textbox source containers overlap')
 if(document.unsupported.some(d=>d.capability==='drawings'&&!selected.has(d.id)))throw new TypeError('Textbox pages cannot omit other drawing restrictions')
 const source_diagnostics=structuredClone(document.unsupported)
 document.unsupported=document.unsupported.filter(d=>!selected.has(d.id))
 return {document,geometry,items,source_diagnostics}
}

export interface NativeDocxTextboxPagesPreviewV2 {
 protocol:typeof DOCX_TEXTBOX_PAGES_PREVIEW_PROTOCOL
 version:2
 fidelity:'approximate'
 read_only:true
 warning:typeof DOCX_TEXTBOX_PAGES_PREVIEW_WARNING
 source_sha256:string
 source_diagnostics:ReturnType<typeof projectNativeDocxTextboxPagesV2>['source_diagnostics']
 body_paint:NativeDocxPagePaintSuccessV1
 textboxes:{textbox_index:number;page_id:string;x_millipoints:number;y_millipoints:number;paint:NativeDocxTextboxShapePaintV1}[]
}

export function decodeNativeDocxTextboxPagesPreviewV2(source:unknown,evidence:unknown,input:unknown,expectedFontSHA256s:readonly string[]):NativeDocxTextboxPagesPreviewV2{
 const projection=projectNativeDocxTextboxPagesV2(source,evidence)
 const value=nativeTextboxGeometryPlainData(input) as NativeDocxTextboxPagesPreviewV2
 const keys=(v:unknown,wanted:string)=>{if(!v||typeof v!=='object'||Array.isArray(v)||Object.keys(v).sort().join(',')!==wanted)throw new TypeError('Invalid textbox pages preview fields')}
 keys(value,'body_paint,fidelity,protocol,read_only,source_diagnostics,source_sha256,textboxes,version,warning')
 if(value.protocol!==DOCX_TEXTBOX_PAGES_PREVIEW_PROTOCOL||value.version!==2||value.fidelity!=='approximate'||value.read_only!==true||value.warning!==DOCX_TEXTBOX_PAGES_PREVIEW_WARNING||value.source_sha256!==nativeTextboxGeometryDigestV1(source)||nativeTextboxGeometryDigestV1(value.source_diagnostics)!==nativeTextboxGeometryDigestV1(projection.source_diagnostics))throw new TypeError('Textbox pages source mismatch')
 if(!Array.isArray(value.textboxes)||value.textboxes.length!==projection.items.length||expectedFontSHA256s.length!==projection.items.length)throw new TypeError('Textbox pages must cover every source rectangle and font')
 const decoded=decodeNativeDocxPagePaintV1(value.body_paint)
 if(!decoded.ok||decoded.value.status!=='painted')throw new TypeError('Textboxes require complete body paint')
 const body=decoded.value,{document}=projection,provenance=body.provenance
 if(provenance.document_id!==document.document_id||provenance.revision!==document.revision||provenance.package_sha256!==document.source.package_sha256||provenance.main_part!==document.source.main_part||provenance.body_story_id!==document.body.id)throw new TypeError('Textbox body source mismatch')
 const textboxes=value.textboxes.map((textbox,ordinal)=>{
  keys(textbox,'page_id,paint,textbox_index,x_millipoints,y_millipoints')
  const {item,index}=projection.items[ordinal]!
  const pages=body.pages.filter(p=>p.lines.some(l=>l.region==='body'&&l.paragraph_id===item.owner.paragraph_id&&l.source_line_ordinal===0)),page=pages[0]
  const paint=decodeNativeDocxTextboxShapePaintV1(source,projection.geometry,index,textbox.paint,expectedFontSHA256s[index]!)
  if(!page)throw new TypeError('Textbox placement has no page')
  const {x,y}=resolveTextboxPosition(document,item,page,paint),half=item.page_anchor!.policy==='page-offset-no-wrap-v1'?paint.line_width_millipoints/2:0
  if(textbox.textbox_index!==index||pages.length!==1||!page||page.kind!=='content'||textbox.page_id!==page.id||textbox.x_millipoints!==x||textbox.y_millipoints!==y||paint.status!=='supported'||x<half||y<half||x+paint.width_millipoints+half>page.width_millipoints||y+paint.height_millipoints+half>page.height_millipoints)throw new TypeError('Textbox placement does not fit its anchor page')
  return {...textbox,paint}
 })
 return {...value,body_paint:body,textboxes}
}
