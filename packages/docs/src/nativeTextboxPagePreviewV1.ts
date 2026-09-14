import {decodeNativeDocxDocument} from './nativeContract.js'
import {decodeNativeDocxPagePaintV1} from './nativePagePaintWireV1.js'
import type {NativeDocxPagePaintSuccessV1} from './nativePagePaintV1.js'
import {decodeNativeDocxTextboxGeometryV1,decodeNativeDocxTextboxShapePaintV1,nativeTextboxGeometryDigestV1,nativeTextboxGeometryPlainData,type NativeDocxTextboxShapePaintV1} from './nativeTextboxGeometryPreviewV1.js'

export const DOCX_TEXTBOX_PAGE_PREVIEW_PROTOCOL='injoffice.docx.textbox-page-preview' as const
export const DOCX_TEXTBOX_PAGE_PREVIEW_WARNING='Read-only approximate textbox composition: one page-relative rectangle anchored before its paragraph text; original drawing restrictions remain.'

/** Source-qualified drawing projection used only by the distinct read-only renderer. */
export function projectNativeDocxTextboxPageV1(source:unknown,evidence:unknown){
 const decoded=decodeNativeDocxDocument(source)
 if(!decoded.ok)throw new TypeError('Invalid textbox page source')
 const document=structuredClone(decoded.value),joined=decodeNativeDocxTextboxGeometryV1(document,evidence)
 if(joined.items.length!==1||joined.omitted_count!==0)throw new TypeError('Textbox pages require one complete rectangle')
 const item=joined.items[0]!,position=item.page_anchor,p=document.body.blocks.find(b=>b.id===item.owner.paragraph_id)?.paragraph
 const diagnostic=document.unsupported.find(d=>d.id===item.owner.diagnostic_id)
 if(!position||position.policy!=='page-offset-no-wrap-v1'||!item.geometry||!p||diagnostic?.code!=='PICTURE_GRAPHIC_REQUIRED'||document.unsupported.some(d=>d.capability==='drawings'&&d.id!==diagnostic.id))throw new TypeError('Unsupported textbox page source')
 if(!position.source_anchor.path.startsWith(p.anchor.path+'/w:r[1]/w:drawing[1]/')||p.runs.some(r=>r.anchor.start_byte<=position.source_anchor.end_byte))throw new TypeError('Textbox must precede all text in its anchor paragraph')
 const source_diagnostics=structuredClone(document.unsupported)
 document.unsupported=document.unsupported.filter(d=>d.id!==diagnostic.id)
 return {document,item,source_diagnostics}
}

export interface NativeDocxTextboxPagePreviewV1 {
 protocol:typeof DOCX_TEXTBOX_PAGE_PREVIEW_PROTOCOL
 version:1
 fidelity:'approximate'
 read_only:true
 warning:typeof DOCX_TEXTBOX_PAGE_PREVIEW_WARNING
 source_sha256:string
 source_diagnostics:ReturnType<typeof projectNativeDocxTextboxPageV1>['source_diagnostics']
 body_paint:NativeDocxPagePaintSuccessV1
 textbox:{page_id:string;x_millipoints:number;y_millipoints:number;paint:NativeDocxTextboxShapePaintV1}
}

/** Browser-safe output validation. Evidence must come from inspection of the
 * current source bytes; this checks joins, not the authenticity of supplied XML hashes. */
export function decodeNativeDocxTextboxPagePreviewV1(source:unknown,evidence:unknown,input:unknown,expectedFontSHA256:string):NativeDocxTextboxPagePreviewV1{
 const projection=projectNativeDocxTextboxPageV1(source,evidence)
 const value=nativeTextboxGeometryPlainData(input) as NativeDocxTextboxPagePreviewV1
 const keys=(v:unknown,wanted:string)=>{if(!v||typeof v!=='object'||Array.isArray(v)||Object.keys(v).sort().join(',')!==wanted)throw new TypeError('Invalid textbox page preview fields')}
 keys(value,'body_paint,fidelity,protocol,read_only,source_diagnostics,source_sha256,textbox,version,warning')
 keys(value.textbox,'page_id,paint,x_millipoints,y_millipoints')
 if(value.protocol!==DOCX_TEXTBOX_PAGE_PREVIEW_PROTOCOL||value.version!==1||value.fidelity!=='approximate'||value.read_only!==true||value.warning!==DOCX_TEXTBOX_PAGE_PREVIEW_WARNING||value.source_sha256!==nativeTextboxGeometryDigestV1(source)||nativeTextboxGeometryDigestV1(value.source_diagnostics)!==nativeTextboxGeometryDigestV1(projection.source_diagnostics))throw new TypeError('Textbox page source mismatch')
 const decoded=decodeNativeDocxPagePaintV1(value.body_paint)
 if(!decoded.ok||decoded.value.status!=='painted')throw new TypeError('Textbox requires complete body paint')
 const body=decoded.value,{document,item}=projection,provenance=body.provenance
 if(provenance.document_id!==document.document_id||provenance.revision!==document.revision||provenance.package_sha256!==document.source.package_sha256||provenance.main_part!==document.source.main_part||provenance.body_story_id!==document.body.id)throw new TypeError('Textbox body source mismatch')
 const pages=body.pages.filter(p=>p.lines.some(l=>l.region==='body'&&l.paragraph_id===item.owner.paragraph_id&&l.source_line_ordinal===0)),page=pages[0]
 const paint=decodeNativeDocxTextboxShapePaintV1(source,evidence,0,value.textbox.paint,expectedFontSHA256),position=item.page_anchor!,x=position.x_emu*10/127,y=position.y_emu*10/127,half=paint.line_width_millipoints/2
 if(pages.length!==1||!page||page.kind!=='content'||value.textbox.page_id!==page.id||value.textbox.x_millipoints!==x||value.textbox.y_millipoints!==y||paint.status!=='supported'||x<half||y<half||x+paint.width_millipoints+half>page.width_millipoints||y+paint.height_millipoints+half>page.height_millipoints)throw new TypeError('Textbox placement does not fit its anchor page')
 return {...value,body_paint:body,textbox:{...value.textbox,paint}}
}
