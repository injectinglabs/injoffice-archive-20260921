/** Node-only atomic composition of source-qualified page textboxes. */
import {prepareNativeDocxPagePaintV1,type NativeDocxPagePaintPrepareInputV1,type NativeDocxHostFontsV1} from './nativePagePaintCompilerV1.js'
import {compileNativeDocxPagePaintV1,type NativeDocxGlyphOutlineProviderV1} from './nativePagePaintV1.js'
import {compileNativeDocxTextboxShapeV1} from './nativeTextboxShapePaintV1.js'
import {nativeTextboxFontDigestV1,nativeTextboxGeometryDigestV1,nativeTextboxGeometryPlainData} from './nativeTextboxGeometryPreviewV1.js'
import {projectNativeDocxTextboxPagesV2,decodeNativeDocxTextboxPagesPreviewV2,DOCX_TEXTBOX_PAGES_PREVIEW_PROTOCOL,DOCX_TEXTBOX_PAGES_PREVIEW_WARNING,type NativeDocxTextboxPagesPreviewV2} from './nativeTextboxPagesPreviewV2.js'

/** Fonts are supplied in evidence item order; composition follows source order.
 * Snapshot all shape inputs before the first asynchronous body compilation. */
export async function renderNativeDocxTextboxPagesPreviewV2(input:NativeDocxPagePaintPrepareInputV1,evidence:unknown,fontBytes:readonly Uint8Array[],outlineProvider:NativeDocxGlyphOutlineProviderV1,runtime?:{fonts?:NativeDocxHostFontsV1}):Promise<NativeDocxTextboxPagesPreviewV2>{
 const source=nativeTextboxGeometryPlainData(input.document),projection=projectNativeDocxTextboxPagesV2(source,evidence)
 if(fontBytes.length!==projection.items.length)throw new TypeError('Every textbox requires its exact font bytes')
 const digests=fontBytes.map(bytes=>nativeTextboxFontDigestV1(bytes))
 const boxes=projection.items.map(({item,index})=>{
  const paint=compileNativeDocxTextboxShapeV1(source,projection.geometry,index,fontBytes[index]!)
  if(paint.status!=='supported')throw new TypeError('Textbox page shape refused: '+paint.reason)
  return {item,index,paint}
 })
 const prepared=await prepareNativeDocxPagePaintV1({...input,document:projection.document},runtime)
 const compiled=await compileNativeDocxPagePaintV1(prepared.page_paint_request,outlineProvider)
 if(!compiled.ok||compiled.value.status!=='painted')throw new TypeError('Textbox page body refused')
 const body=compiled.value
 const textboxes=boxes.map(({item,index,paint})=>{
  const page=body.pages.find(p=>p.lines.some(l=>l.region==='body'&&l.paragraph_id===item.owner.paragraph_id&&l.source_line_ordinal===0))
  if(!page)throw new TypeError('Textbox anchor paragraph has no page')
  return {textbox_index:index,page_id:page.id,x_millipoints:item.page_anchor!.x_emu*10/127,y_millipoints:item.page_anchor!.y_emu*10/127,paint}
 })
 return decodeNativeDocxTextboxPagesPreviewV2(source,projection.geometry,{protocol:DOCX_TEXTBOX_PAGES_PREVIEW_PROTOCOL,version:2,fidelity:'approximate',read_only:true,warning:DOCX_TEXTBOX_PAGES_PREVIEW_WARNING,source_sha256:nativeTextboxGeometryDigestV1(source),source_diagnostics:projection.source_diagnostics,body_paint:body,textboxes},digests)
}
