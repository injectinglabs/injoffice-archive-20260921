import type {NativePptxTextLayout,RenderRect,RenderTextBodyNode} from './types.js'

const floor=(n:bigint,d:bigint)=>n/d-(n<0n&&n%d!==0n?1n:0n)
function scale(value:number,size:number,units:number,upper:boolean):number {
 const n=BigInt(value)*BigInt(size)*127n,d=BigInt(units)*10n
 return Number(upper?-floor(-n,d):floor(n,d))
}
/** Supplied outline control hulls bound overflow and marks before an arbitrary
 * affine is painted. Identity/digest validation mirrors the font provider join. */
export async function sourceTextBounds(body:RenderTextBodyNode,extents:NativePptxTextLayout['glyphExtents']):Promise<RenderRect> {
 let minX=body.bounds.x,minY=body.bounds.y,maxX=minX+body.bounds.cx,maxY=minY+body.bounds.cy
 const include=(x:number,y:number)=>{if(!Number.isSafeInteger(x)||!Number.isSafeInteger(y))throw new RangeError('Source text hull exceeds integer precision');minX=Math.min(minX,x);minY=Math.min(minY,y);maxX=Math.max(maxX,x);maxY=Math.max(maxY,y)}
 for(const paragraph of body.paragraphs){
  include(paragraph.x,paragraph.y);include(paragraph.x+paragraph.widthEmu,paragraph.y+paragraph.heightEmu)
  for(const run of paragraph.marker?[paragraph.marker,...paragraph.runs]:paragraph.runs){
   if(run.status==='refused')continue
   if(run.glyphs.length&&(!extents||!run.faceId||!run.contentDigest))throw new RangeError('Arbitrary source transforms require supplied glyph control hulls')
   for(const glyph of run.glyphs){
    const value=await extents!({faceId:run.faceId!,contentDigest:run.contentDigest!,glyphId:glyph.glyphId})
    if(value.faceId!==run.faceId||value.contentDigest!==run.contentDigest||value.glyphId!==glyph.glyphId||!Number.isSafeInteger(value.unitsPerEm)||value.unitsPerEm<1||value.unitsPerEm>1000000000)throw new RangeError('Source text outline identity or design grid mismatch')
    if(value.bounds===null)continue
    const b=value.bounds
    if([b.xMin,b.yMin,b.xMax,b.yMax].some(v=>!Number.isSafeInteger(v)||Math.abs(v)>1000000000)||b.xMin>b.xMax||b.yMin>b.yMax)throw new RangeError('Invalid source text outline hull')
    include(run.x+glyph.xEmu+scale(b.xMin,run.fontSizeMilliPoints,value.unitsPerEm,false),run.baselineY+glyph.yEmu-scale(b.yMax,run.fontSizeMilliPoints,value.unitsPerEm,true))
    include(run.x+glyph.xEmu+scale(b.xMax,run.fontSizeMilliPoints,value.unitsPerEm,true),run.baselineY+glyph.yEmu-scale(b.yMin,run.fontSizeMilliPoints,value.unitsPerEm,false))
   }
  }
 }
 if(!Number.isSafeInteger(maxX-minX)||!Number.isSafeInteger(maxY-minY))throw new RangeError('Source text hull extent exceeds integer precision')
 return {x:minX,y:minY,cx:maxX-minX,cy:maxY-minY}
}
