import {decodeNativeDocxDocument,decodeNativeDocxTextboxGeometryV1,decodeNativeDocxTextboxPagePreviewV1,decodeNativeDocxTextboxPagesPreviewV2,type NativeDocxTextboxPagePreviewV1} from '@injoffice/docs/native-docx'
import {decodeNativeDOCXFontInventoryV1} from '../../../../packages/docs/src/nativeFontInventoryV1'

/** Current package bytes and every embedded font independently join the paint. */
function textboxPageSource(value:Record<string,unknown>,packageDigest:string){
 if(Object.keys(value).sort().join(',')!=='document,evidence,font_inventory_json,preview'||typeof value.font_inventory_json!=='string')throw new TypeError('Invalid textbox preview response')
 const document=decodeNativeDocxDocument(value.document)
 if(!document.ok||document.value.source.package_sha256!==packageDigest)throw new TypeError('Textbox pages do not match the currently opened document')
 const source=document.value,evidence=decodeNativeDocxTextboxGeometryV1(source,value.evidence)
 const inventory=decodeNativeDOCXFontInventoryV1(value.font_inventory_json)
 if(inventory.document_id!==source.document_id||inventory.revision!==source.revision||inventory.package_sha256!==packageDigest||inventory.main_part!==source.source.main_part||!evidence.items.length||evidence.items.some(item=>inventory.main_sha256!==item.owner.part_sha256))throw new TypeError('Textbox font inventory source mismatch')
 const digests=evidence.items.map(item=>{
  const faces=inventory.families.flatMap(f=>f.faces).filter(f=>f.family===item.geometry?.font_family&&f.weight===400&&f.style==='normal'&&(f.stretch??100)===100&&f.source.face_slot==='embedRegular')
  const face=faces[0]
  if(faces.length!==1||!face||!source.passthrough_parts.some(p=>p.part_name===face.source.asset_part&&p.sha256===face.source.stored_sha256&&p.byte_length===face.source.stored_byte_length))throw new TypeError('Textbox preview requires an exact source-embedded regular font')
  return face.source.content_sha256
 })
 return {source,evidence,digests}
}

export function decodeTextboxPageResponse(value:Record<string,unknown>,packageDigest:string):NativeDocxTextboxPagePreviewV1{
 const {source,evidence,digests}=textboxPageSource(value,packageDigest)
 return decodeNativeDocxTextboxPagePreviewV1(source,evidence,value.preview,digests[0]!)
}

type TextboxPages=Pick<NativeDocxTextboxPagePreviewV1,'body_paint'|'source_diagnostics'> & {warning:string;textboxes:NativeDocxTextboxPagePreviewV1['textbox'][]}
export function decodeTextboxesPageResponse(value:Record<string,unknown>,packageDigest:string):TextboxPages{
 const {source,evidence,digests}=textboxPageSource(value,packageDigest)
 if(value.preview&&typeof value.preview==='object'&&(value.preview as {version?:unknown}).version===2)return decodeNativeDocxTextboxPagesPreviewV2(source,evidence,value.preview,digests)
 const legacy=decodeNativeDocxTextboxPagePreviewV1(source,evidence,value.preview,digests[0]!)
 return {...legacy,textboxes:[legacy.textbox]}
}

export function NativeDocxTextboxOnPage({textbox,pageID}:{textbox:NativeDocxTextboxPagePreviewV1['textbox'];pageID:string}){
 if(textbox.page_id!==pageID)return null
 const p=textbox.paint,color=(rgb:string)=>rgb==='none'?'none':'#'+rgb
 return <g data-native-textbox="true" data-native-textbox-id={p.diagnostic_id} transform={`translate(${textbox.x_millipoints} ${textbox.y_millipoints})`}>
  <rect width={p.width_millipoints} height={p.height_millipoints} fill={color(p.fill_rgb)} stroke={color(p.line_rgb)} strokeWidth={p.line_width_millipoints}/>
  {p.paths.map((path,i)=><path key={i} d={path} fill={color(p.text_rgb)} fillRule="nonzero"/>)}
 </g>
}
