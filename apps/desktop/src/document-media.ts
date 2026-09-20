import type {NativeDocxDocumentV1,NativeDocxDrawingV1} from '../../../packages/docs/src/nativeContract'
export interface DocumentMediaReader {readMedia(bytes:Uint8Array,document:NativeDocxDocumentV1,id:string):Promise<Uint8Array>}
export type DocumentImageCache=Map<string,{url:string;size:number}>
const MAX_BYTES=8*1024*1024,MAX_PIXELS=32*1024*1024
/**
 * The data-URL prefix is chosen from this fixed table, never assembled from the document's declared
 * content type. Splicing an untrusted `content_type` into a URL is the classic reinterpretation sink
 * (CodeQL js/xss-through-dom); a lookup keyed by it returns only a literal we wrote.
 */
const RASTER_DATA_URL_PREFIX:Readonly<Record<string,string>>=Object.freeze({'image/png':'data:image/png;base64,','image/jpeg':'data:image/jpeg;base64,'})
export function rasterDataUrlPrefix(contentType:string|undefined):string|undefined {
  return contentType!==undefined&&Object.prototype.hasOwnProperty.call(RASTER_DATA_URL_PREFIX,contentType)?RASTER_DATA_URL_PREFIX[contentType]:undefined
}
/** A preview may only paint an <img> whose src is one of our own raster data URLs. */
export function isRasterDataUrl(url:string|undefined):url is string {
  return typeof url==='string'&&Object.values(RASTER_DATA_URL_PREFIX).some(prefix=>url.startsWith(prefix))
}
export function documentDrawings(document:NativeDocxDocumentV1):NativeDocxDrawingV1[] {
  return [document.body,...document.headers,...document.footers,...document.notes,...document.comment_stories].flatMap(story=>story.blocks.flatMap(block=>block.paragraph?[block.paragraph]:block.table?.rows.flatMap(row=>row.cells.flatMap(cell=>cell.paragraphs))??[]).flatMap(paragraph=>paragraph.runs.flatMap(run=>run.drawing?[run.drawing]:[])))
}
/** Only source-qualified inline rasters are shown. Unsupported placement remains a placeholder. */
export async function loadDocumentImages(reader:DocumentMediaReader,bytes:Uint8Array,document:NativeDocxDocumentV1,cache:DocumentImageCache):Promise<{images:Record<string,string>;notice:string}> {
  const images:Record<string,string>=Object.create(null)
  let pixels=0,total=0,count=0,omitted=0
  for(const drawing of documentDrawings(document)) {
    const raster=drawing.raster
    const prefix=rasterDataUrlPrefix(drawing.content_type)
    if(!raster||drawing.placement!=='inline'||!prefix||drawing.width_emu<=0||drawing.height_emu<=0){omitted++;continue}
    const area=raster.pixel_width*raster.pixel_height
    if(++count>64||pixels+area>MAX_PIXELS){omitted++;continue}
    const key=`${drawing.content_type}:${raster.sha256}`
    try {
      let asset=cache.get(key)
      if(!asset){
        const data=await reader.readMedia(bytes,document,drawing.id)
        if(data.length>2*1024*1024||total+data.length>MAX_BYTES){omitted++;continue}
        let binary='';for(let i=0;i<data.length;i+=8192)binary+=String.fromCharCode(...data.subarray(i,i+8192))
        asset={url:prefix+btoa(binary),size:data.length}
        while(cache.size&&[...cache.values()].reduce((sum,value)=>sum+value.size,0)+asset.size>MAX_BYTES)cache.delete(cache.keys().next().value!)
        cache.set(key,asset)
      }
      if(total+asset.size>MAX_BYTES){omitted++;continue}
      total+=asset.size;pixels+=area;images[drawing.id]=asset.url
    }catch{omitted++}
  }
  return {images,notice:omitted?`${omitted} image${omitted===1?' is':'s are'} preserved but not shown because placement, image features, or preview limits are unsupported.`:''}
}
