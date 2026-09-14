import {sourceAffine,addSourceAffineRationals,subtractSourceAffineRationals,type AffineRational,type SourceAffineFrame,composeSourceAffines,convertSourceAffine,decodeSourceAffine,qualifySourceAffinePoint,sourceAffineRational,SourceAffineBudget,type QualifiedSourceAffine} from './sourceAffine.js'
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


export interface SourceGraphicFrameLayout {
 readonly policy:'office-graphic-frame-anchor-v1'
 readonly width:AffineRational
 readonly height:AffineRational
 /** Translation only: glyph metrics and intrinsic table tracks are not scaled. */
 readonly origin:QualifiedSourceAffine
}

/** Named implementation policy, qualified against the retained Office matrix.
 * Direct graphic-frame orientation is not painted. Ancestors map its center;
 * the leaf's raw rotation quadrant chooses the physical anchor's scale axes.
 * Width/height remain rational: callers must not round these layout thresholds.
 * This helper alone does not admit source or authorize mutations. */
export function sourceGraphicFrameLayout(leaf:SourceAffineFrame,parents:readonly {frame:SourceAffineFrame;child:Pick<SourceAffineFrame,'x'|'y'|'cx'|'cy'>}[],budget:SourceAffineBudget):SourceGraphicFrameLayout {
 if(!Array.isArray(parents)||parents.length>=64)throw new RangeError('Graphic-frame hierarchy depth exceeded')
 sourceAffine(leaf,undefined,budget) // Reuse complete frame/angle/boolean validation.
 const zero=sourceAffineRational(0n),one=sourceAffineRational(1n)
 const integer=(n:number)=>sourceAffineRational(BigInt(n))
 const half=(n:AffineRational)=>sourceAffineRational(n.numerator,n.denominator*2n)
 const multiply=(a:AffineRational,b:AffineRational)=>{budget.charge(16);return sourceAffineRational(a.numerator*b.numerator,a.denominator*b.denominator)}
 let center:QualifiedSourceAffine={values:[one,zero,zero,one,addSourceAffineRationals(integer(leaf.x),sourceAffineRational(BigInt(leaf.cx),2n)),addSourceAffineRationals(integer(leaf.y),sourceAffineRational(BigInt(leaf.cy),2n))],errors:[zero,zero,zero,zero,zero,zero],depth:1}
 let sx=one,sy=one
 for(let i=0;i<parents.length;i++){
  const parent=parents[i]
  if(!parent)throw new RangeError('Missing graphic-frame ancestor')
  center=composeSourceAffines(sourceAffine(parent.frame,parent.child,budget),center,budget)
  sx=multiply(sx,sourceAffineRational(BigInt(parent.frame.cx),BigInt(parent.child.cx)))
  sy=multiply(sy,sourceAffineRational(BigInt(parent.frame.cy),BigInt(parent.child.cy)))
 }
 const angle=leaf.rotation??0,swap=angle>=2700000&&angle<8100000||angle>=13500000&&angle<18900000
 const width=multiply(integer(leaf.cx),swap?sy:sx),height=multiply(integer(leaf.cy),swap?sx:sy)
 budget.charge(32)
 const origin:QualifiedSourceAffine={values:[one,zero,zero,one,subtractSourceAffineRationals(center.values[4],half(width)),subtractSourceAffineRationals(center.values[5],half(height))],errors:[zero,zero,zero,zero,center.errors[4],center.errors[5]],depth:center.depth}
 return {policy:'office-graphic-frame-anchor-v1',width,height,origin}
}
