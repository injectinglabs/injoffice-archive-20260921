import {composeSourceAffines,SourceAffineBudget,type QualifiedSourceAffine} from './sourceAffine.js'
import {sourceTextBounds} from './sourceTextBounds.js'
import {sourceTableTextClip} from './sourceTableTextClip.js'
import {qualifySourceGraphicFrameSpans,sourceGraphicFramePaintTransform,SOURCE_GRAPHIC_FRAME_MAX_SPANS,type SourceGraphicFrameSpan} from './sourceGraphicFramePolicy.js'
import type {NativePptxTextLayout,RenderRect,RenderTextBodyNode} from './types.js'

const identity={aPpm:1000000,bPpm:0,cPpm:0,dPpm:1000000,txEmu:0,tyEmu:0} as const
/** Mirror paintTextBody's exact transform order, including the body-refusal
 * short circuit. Paragraph-local fractional placement must be applied to both
 * supplied glyphs and refused-run placeholders. The caller supplies the actual
 * cell world (including cell translation), and keeps backgrounds/borders outside
 * this text-only clip. Source/body admission is deliberately unchanged here. */
export async function sourceTableTextBounds(body:RenderTextBodyNode,cellWidth:number,world:QualifiedSourceAffine,extents:NativePptxTextLayout['glyphExtents'],maxCoordinateEmu:number,budget:SourceAffineBudget):Promise<{readonly spans:readonly SourceGraphicFrameSpan[];readonly horizontalClip?:RenderRect}> {
 if(!Number.isSafeInteger(cellWidth)||cellWidth<=0||cellWidth>maxCoordinateEmu||!Array.isArray(body.paragraphs)||body.paragraphs.length>SOURCE_GRAPHIC_FRAME_MAX_SPANS)throw new RangeError('Invalid bounded table text body')
 const unchanged=sourceGraphicFramePaintTransform(identity,budget)
 const spans:SourceGraphicFrameSpan[]=[]
 if(body.status==='refused')spans.push({bounds:body.bounds,transform:unchanged})
 else{
  const orientation=sourceGraphicFramePaintTransform(body.orientationTransform??identity,budget)
  const vertical=sourceGraphicFramePaintTransform(body.transform??identity,budget)
  const local=composeSourceAffines(orientation,vertical,budget)
  for(let i=0;i<body.paragraphs.length;i++){
   const paragraph=body.paragraphs[i];if(!paragraph)throw new RangeError('Missing table paragraph')
   const bounds=await sourceTextBounds({...body,bounds:{x:0,y:0,cx:0,cy:0},paragraphs:[paragraph]},extents)
   const transform=composeSourceAffines(local,sourceGraphicFramePaintTransform(paragraph.transform??identity,budget),budget)
   spans.push({bounds,transform})
  }
  if(!spans.length)spans.push({bounds:body.bounds,transform:local})
 }
 qualifySourceGraphicFrameSpans(world,spans,maxCoordinateEmu,budget)
 if(body.horizontalOverflow!=='clip')return {spans}
 const horizontalClip=sourceTableTextClip(cellWidth,spans,maxCoordinateEmu,budget)
 qualifySourceGraphicFrameSpans(world,[{bounds:horizontalClip,transform:unchanged}],maxCoordinateEmu,budget)
 return {spans,horizontalClip}
}
