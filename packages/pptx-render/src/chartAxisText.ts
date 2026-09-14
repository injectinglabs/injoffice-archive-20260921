import type {NativePptxTextLayout,RenderRect,RenderTextBodyNode} from './types.js'

function floor(n:bigint,d:bigint):bigint{return n/d-(n<0n&&n%d!==0n?1n:0n)}
function scale(bound:number,size:number,units:number,upper:boolean):number{
 const n=BigInt(bound)*BigInt(size)*127n,d=BigInt(units)*10n
 return Number(upper?-floor(-n,d):floor(n,d))
}
/** Bounds the actual positioned outlines conservatively, including bearings
 * and mark offsets, alongside typographic advances. Curve control hulls may
 * reserve more space than the true ink; no exact curve-extrema claim is made. */
export async function measureChartAxisText(body:RenderTextBodyNode,extents:NativePptxTextLayout['glyphExtents']):Promise<RenderRect>{
 if(!extents||body.status!=='laidOut'||!['native','deterministicNative'].includes(body.fidelity)||body.paragraphs.length!==1)throw new RangeError('exact supplied-font axis text is unavailable')
 const paragraph=body.paragraphs[0]!
 if(paragraph.runs.length!==1||paragraph.runs[0]!.status!=='shaped'||paragraph.runs[0]!.fontSelection?.resolution!=='exact')throw new RangeError('axis labels require one exact source-font run')
 let xMin=paragraph.x,yMin=paragraph.y,xMax=paragraph.x+paragraph.widthEmu,yMax=paragraph.y+paragraph.heightEmu
 const run=paragraph.runs[0]!
 if(!run.faceId||!run.contentDigest)throw new RangeError('axis font identity is missing')
 if(run.glyphs.length>32768)throw new RangeError('axis glyph budget exceeded')
 for(const glyph of run.glyphs){
  const value=await extents({faceId:run.faceId,contentDigest:run.contentDigest,glyphId:glyph.glyphId})
  if(value.faceId!==run.faceId||value.contentDigest!==run.contentDigest||value.glyphId!==glyph.glyphId||!Number.isSafeInteger(value.unitsPerEm)||value.unitsPerEm<1||value.unitsPerEm>1000000000)throw new RangeError('axis outline identity or design grid mismatch')
  if(value.bounds===null)continue
  const b=value.bounds
  if([b.xMin,b.yMin,b.xMax,b.yMax].some(v=>!Number.isSafeInteger(v)||Math.abs(v)>1000000000)||b.xMin>b.xMax||b.yMin>b.yMax)throw new RangeError('invalid axis outline bounds')
  xMin=Math.min(xMin,run.x+glyph.xEmu+scale(b.xMin,run.fontSizeMilliPoints,value.unitsPerEm,false))
  xMax=Math.max(xMax,run.x+glyph.xEmu+scale(b.xMax,run.fontSizeMilliPoints,value.unitsPerEm,true))
  yMin=Math.min(yMin,run.baselineY+glyph.yEmu-scale(b.yMax,run.fontSizeMilliPoints,value.unitsPerEm,true))
  yMax=Math.max(yMax,run.baselineY+glyph.yEmu-scale(b.yMin,run.fontSizeMilliPoints,value.unitsPerEm,false))
 }
 if([xMin,yMin,xMax,yMax,xMax-xMin,yMax-yMin].some(v=>!Number.isSafeInteger(v)||Math.abs(v)>281474976710655)||xMax<=xMin||yMax<=yMin)throw new RangeError('axis label measurement budget exceeded')
 return {x:xMin,y:yMin,cx:xMax-xMin,cy:yMax-yMin}
}
