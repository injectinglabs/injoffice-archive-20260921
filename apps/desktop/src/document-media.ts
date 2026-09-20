import type {NativeDocxDocumentV1,NativeDocxDrawingV1} from '../../../packages/docs/src/nativeContract'
export interface DocumentMediaReader {readMedia(bytes:Uint8Array,document:NativeDocxDocumentV1,id:string):Promise<Uint8Array>}
export type DocumentImageCache=Map<string,{url:string;size:number}>
const MAX_BYTES=8*1024*1024,MAX_PIXELS=32*1024*1024
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
/** Only source-qualified inline rasters are shown. Unsupported placement remains a placeholder. */
export async function loadDocumentImages(reader:DocumentMediaReader,bytes:Uint8Array,document:NativeDocxDocumentV1,cache:DocumentImageCache):Promise<{images:Record<string,string>;notice:string}> {
  const images:Record<string,string>=Object.create(null)
  let pixels=0,total=0,count=0,omitted=0
  for(const drawing of documentDrawings(document)) {
    const raster=drawing.raster
    const type=rasterContentType(drawing.content_type)
    if(!raster||drawing.placement!=='inline'||!type||drawing.width_emu<=0||drawing.height_emu<=0){omitted++;continue}
    const area=raster.pixel_width*raster.pixel_height
    if(++count>64||pixels+area>MAX_PIXELS){omitted++;continue}
    const key=`${drawing.content_type}:${raster.sha256}`
    try {
      let asset=cache.get(key)
      if(!asset){
        const data=await reader.readMedia(bytes,document,drawing.id)
        if(data.length>2*1024*1024||total+data.length>MAX_BYTES){omitted++;continue}
        asset={url:URL.createObjectURL(new Blob([data],{type})),size:data.length}
        while(cache.size&&[...cache.values()].reduce((sum,value)=>sum+value.size,0)+asset.size>MAX_BYTES)evict(cache,cache.keys().next().value!)
        cache.set(key,asset)
      }
      if(total+asset.size>MAX_BYTES){omitted++;continue}
      total+=asset.size;pixels+=area;images[drawing.id]=asset.url
    }catch{omitted++}
  }
  return {images,notice:omitted?`${omitted} image${omitted===1?' is':'s are'} preserved but not shown because placement, image features, or preview limits are unsupported.`:''}
}
