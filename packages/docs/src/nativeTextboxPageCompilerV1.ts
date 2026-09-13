/** Node-only page composition over source-qualified geometry and actual fonts. */
import {prepareNativeDocxPagePaintV1,type NativeDocxPagePaintPrepareInputV1,type NativeDocxHostFontsV1} from './nativePagePaintCompilerV1.js'
import {compileNativeDocxPagePaintV1,type NativeDocxGlyphOutlineProviderV1} from './nativePagePaintV1.js'
import {compileNativeDocxTextboxShapeV1} from './nativeTextboxShapePaintV1.js'
import {nativeTextboxFontDigestV1,nativeTextboxGeometryDigestV1,nativeTextboxGeometryPlainData} from './nativeTextboxGeometryPreviewV1.js'
import {projectNativeDocxTextboxPageV1,decodeNativeDocxTextboxPagePreviewV1,DOCX_TEXTBOX_PAGE_PREVIEW_PROTOCOL,DOCX_TEXTBOX_PAGE_PREVIEW_WARNING,type NativeDocxTextboxPagePreviewV1} from './nativeTextboxPagePreviewV1.js'

/** Atomically compose a single rectangle at the page chosen by body pagination.
 * Unsupported body, shape, font, or placement throws without returning partial pages. */
export async function renderNativeDocxTextboxPagePreviewV1(input:NativeDocxPagePaintPrepareInputV1,evidence:unknown,fontBytes:Uint8Array,outlineProvider:NativeDocxGlyphOutlineProviderV1,runtime?:{fonts?:NativeDocxHostFontsV1}):Promise<NativeDocxTextboxPagePreviewV1>{
 const source=nativeTextboxGeometryPlainData(input.document)
 const projection=projectNativeDocxTextboxPageV1(source,evidence)
 const geometry={items:[projection.item],omitted_count:0}
 const paint=compileNativeDocxTextboxShapeV1(source,geometry,0,fontBytes)
 if(paint.status!=='supported')throw new TypeError('Textbox page shape refused: '+paint.reason)
 const prepared=await prepareNativeDocxPagePaintV1({...input,document:projection.document},runtime)
 const compiled=await compileNativeDocxPagePaintV1(prepared.page_paint_request,outlineProvider)
 if(!compiled.ok||compiled.value.status!=='painted')throw new TypeError('Textbox page body refused')
 const page=compiled.value.pages.find(p=>p.lines.some(l=>l.region==='body'&&l.paragraph_id===projection.item.owner.paragraph_id&&l.source_line_ordinal===0))
 if(!page)throw new TypeError('Textbox anchor paragraph has no page')
 const result:NativeDocxTextboxPagePreviewV1={protocol:DOCX_TEXTBOX_PAGE_PREVIEW_PROTOCOL,version:1,fidelity:'approximate',read_only:true,warning:DOCX_TEXTBOX_PAGE_PREVIEW_WARNING,source_sha256:nativeTextboxGeometryDigestV1(source),source_diagnostics:projection.source_diagnostics,body_paint:compiled.value,textbox:{page_id:page.id,x_millipoints:projection.item.page_anchor!.x_emu*10/127,y_millipoints:projection.item.page_anchor!.y_emu*10/127,paint}}
 return decodeNativeDocxTextboxPagePreviewV1(source,geometry,result,nativeTextboxFontDigestV1(fontBytes))
}
