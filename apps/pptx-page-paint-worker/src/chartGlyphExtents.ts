import type {NativePptxGlyphExtents,NativePptxGlyphExtentsRequest} from '@injoffice/pptx-render'
import type {FontResource} from '@injoffice/font-metrics/layout'
import type {HarfBuzzOutlineCommandV1} from '@injoffice/font-metrics/harfbuzz'
interface OutlineProvider {outline(glyphId:number):{units_per_em:number;path:HarfBuzzOutlineCommandV1[]}}
/** Reuses the same digest-bound provider used by glyph paint; no new font load. */
export function chartGlyphExtents(resources:ReadonlyMap<string,FontResource>,outlines:ReadonlyMap<string,OutlineProvider>){
 return (request:NativePptxGlyphExtentsRequest):NativePptxGlyphExtents=>{
  const provider=outlines.get(request.faceId),resource=resources.get(request.faceId)
  if(!provider||resource?.face.contentDigest!==request.contentDigest)throw new Error('Chart glyph outline font identity mismatch')
  const outline=provider.outline(request.glyphId)
  let xMin=Infinity,yMin=Infinity,xMax=-Infinity,yMax=-Infinity
  const point=(x:number,y:number)=>{xMin=Math.min(xMin,x);xMax=Math.max(xMax,x);yMin=Math.min(yMin,y);yMax=Math.max(yMax,y)}
  for(const p of outline.path){
   if(p.kind==='close_path')continue
   point(p.x,p.y)
   if(p.kind==='quadratic_to'){point(p.control_x,p.control_y)}
   if(p.kind==='cubic_to'){point(p.control_1_x,p.control_1_y);point(p.control_2_x,p.control_2_y)}
  }
  const bounds=xMin===Infinity?null:{xMin,xMax,yMin,yMax}
  return {...request,unitsPerEm:outline.units_per_em,bounds}
 }
}
