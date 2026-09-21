import {extractDocxPreviewImages} from '../../playground/src/docxPreviewImages'
import type {NativeDocxDocumentV1,NativeDocxDrawingV1} from '../../../packages/docs/src/nativeContract'
export type DocumentImageCache=Map<string,{url:string;size:number}>
/**
 * Preview images are blob: URLs minted from the media bytes with a content type taken from this
 * fixed table, never from the document's declared string. The browser types the blob itself, the
 * <img> receives an opaque URL, and no string derived from document bytes is ever spliced into an
 * attribute (CodeQL js/xss-through-dom, alert #48/#49).
 */
const RASTER_CONTENT_TYPES:Readonly<Record<string,'image/png'|'image/jpeg'>>=Object.freeze({'image/png':'image/png','image/jpeg':'image/jpeg'})
export function rasterContentType(contentType:string|undefined):'image/png'|'image/jpeg'|undefined {
  return contentType!==undefined&&Object.prototype.hasOwnProperty.call(RASTER_CONTENT_TYPES,contentType)?RASTER_CONTENT_TYPES[contentType]:undefined
}
/** A preview may only paint an <img> whose src is a blob URL this module minted. */
export function isPreviewImageUrl(url:string|undefined):url is string {
  return typeof url==='string'&&url.startsWith('blob:')
}
function evict(cache:DocumentImageCache,key:string):void {
  const asset=cache.get(key); cache.delete(key)
  if(asset)try{URL.revokeObjectURL(asset.url)}catch{/* already revoked or unsupported host */}
}
/** Release every minted URL (call when the document is closed). */
export function releaseDocumentImages(cache:DocumentImageCache):void { for(const key of [...cache.keys()])evict(cache,key) }
export function documentDrawings(document:NativeDocxDocumentV1):NativeDocxDrawingV1[] {
  return [document.body,...document.headers,...document.footers,...document.notes,...document.comment_stories].flatMap(story=>story.blocks.flatMap(block=>block.paragraph?[block.paragraph]:block.table?.rows.flatMap(row=>row.cells.flatMap(cell=>cell.paragraphs))??[]).flatMap(paragraph=>paragraph.runs.flatMap(run=>run.drawing?[run.drawing]:[])))
}
/** Read only media whose package and part hashes match native extraction. */
export async function loadSourceDocumentImages(bytes: Uint8Array, document: NativeDocxDocumentV1, cache: DocumentImageCache) {
  const images: Record<string, string> = Object.create(null)
  const drawings = documentDrawings(document)
  if (!drawings.length) return { images, notice: '' }
  const media = await extractDocxPreviewImages(bytes, document)
  let omitted = 0, total = 0
  for (const drawing of drawings) {
    const image = drawing.media_part ? media.get(drawing.media_part) : undefined
    const type = rasterContentType(image?.mime)
    const part = document.passthrough_parts.find(part => part.part_name === drawing.media_part)
    if (!image || !type || !part || drawing.placement !== 'inline' || drawing.source_crop || drawing.flip_horizontal || drawing.flip_vertical || drawing.rotation_degrees || drawing.rotation_60000ths || !(drawing.width_emu! > 0) || !(drawing.height_emu! > 0) || image.bytes.length > 2 * 1024 * 1024 || total + image.bytes.length > 8 * 1024 * 1024) { omitted++; continue }
    const key = `${type}:${part.sha256}`
    let asset = cache.get(key)
    if (!asset) {
      while (cache.size && [...cache.values()].reduce((sum, value) => sum + value.size, 0) + image.bytes.length > 8 * 1024 * 1024) {
        const oldest = cache.keys().next().value!
        URL.revokeObjectURL(cache.get(oldest)!.url)
        cache.delete(oldest)
      }
      asset = { url: URL.createObjectURL(new Blob([Uint8Array.from(image.bytes).buffer], { type })), size: image.bytes.length }
      cache.set(key, asset)
    }
    total += asset.size
    images[drawing.id] = asset.url
  }
  return { images, notice: omitted ? `${omitted} image${omitted === 1 ? ' is' : 's are'} preserved but not shown because placement, image features, or preview limits are unsupported.` : '' }
}
