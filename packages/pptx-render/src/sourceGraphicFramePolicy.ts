import {composeSourceAffines,convertSourceAffine,decodeSourceAffine,qualifySourceAffinePoint,sourceAffineRational,SourceAffineBudget,type QualifiedSourceAffine} from './sourceAffine.js'
import {renderTransformMatrix} from './sourceRenderTransform.js'
import type {RenderRect,RenderTransform} from './types.js'

export const SOURCE_GRAPHIC_FRAME_MAX_SPANS=8192
export interface SourceGraphicFrameSpan {
 readonly bounds:RenderRect
 readonly transform:QualifiedSourceAffine
}

/** Internal adapter for already compiled paint transforms. Reuse the public
 * transport decoder; never reconstruct rational coefficients from binary64. */
export function sourceGraphicFramePaintTransform(transform:RenderTransform,budget:SourceAffineBudget):QualifiedSourceAffine {
 renderTransformMatrix(transform,budget)
 if(transform.sourceAffine)return convertSourceAffine(decodeSourceAffine(transform.sourceAffine,budget),budget).qualified
 const raw=[transform.aPpm,transform.bPpm,transform.cPpm,transform.dPpm,transform.txEmu,transform.tyEmu]
 const values=raw.map((v,i)=>sourceAffineRational(BigInt(v),i<4?1000000n:1n)) as unknown as QualifiedSourceAffine['values']
 const zero=sourceAffineRational(0n)
 return convertSourceAffine({values,errors:[zero,zero,zero,zero,zero,zero],depth:1},budget).qualified
}

/** Qualify actual complete paint/control hulls, not merely the graphic-frame
 * rectangle. The caller must include axis labels/ticks/stroke envelopes and all
 * table glyphs/placeholders. This helper does not choose a text reflection policy
 * or authorize source admission; those remain independently reference-qualified. */
export function qualifySourceGraphicFrameSpans(world:QualifiedSourceAffine,spans:readonly SourceGraphicFrameSpan[],maxCoordinateEmu:number,budget:SourceAffineBudget):void {
 if(!Array.isArray(spans)||spans.length===0||spans.length>SOURCE_GRAPHIC_FRAME_MAX_SPANS)throw new RangeError('Graphic-frame paint span budget exceeded')
 for(let i=0;i<spans.length;i++){
  budget.charge(16)
  const span=spans[i]
  if(!span)throw new RangeError('Missing graphic-frame paint span')
  const {x,y,cx,cy}=span.bounds
  if([x,y,cx,cy,x+cx,y+cy].some(v=>!Number.isSafeInteger(v))||cx<0||cy<0)throw new RangeError('Invalid graphic-frame paint hull')
  // Painting converts each transform before concatenation. Carry both local
  // conversion allowances through the final world operation before qualifying.
  const painted=composeSourceAffines(convertSourceAffine(world,budget).qualified,convertSourceAffine(span.transform,budget).qualified,budget)
  for(const [px,py] of [[x,y],[x+cx,y],[x+cx,y+cy],[x,y+cy]] as const)qualifySourceAffinePoint(painted,px,py,maxCoordinateEmu,budget)
 }
}
