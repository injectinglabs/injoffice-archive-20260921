import {sourceAffineRational,composeSourceAffines,convertSourceAffine,qualifySourceAffinePoint,SourceAffineBudget,SOURCE_AFFINE_LIMITS,type AffineRational,type QualifiedSourceAffine} from './sourceAffine.js'
import type {ExactTextArea} from './sourceTextOrientation.js'

const zero=sourceAffineRational(0n),one=sourceAffineRational(1n)
function checked(value:AffineRational,budget:SourceAffineBudget):AffineRational {
  budget.charge(16)
  if(!value||typeof value.numerator!=='bigint'||typeof value.denominator!=='bigint'||value.denominator<=0n)throw new RangeError('Invalid exact text placement')
  if((value.numerator<0n?-value.numerator:value.numerator).toString(2).length>SOURCE_AFFINE_LIMITS.rationalBits||value.denominator.toString(2).length>SOURCE_AFFINE_LIMITS.rationalBits)throw new RangeError('Text placement precision budget exceeded')
  return sourceAffineRational(value.numerator,value.denominator)
}
function diagonal(sx:AffineRational,sy:AffineRational,x:AffineRational,y:AffineRational):QualifiedSourceAffine {
  return {values:[sx,zero,zero,sy,x,y],errors:[zero,zero,zero,zero,zero,zero],depth:1}
}

/** Fractional placement belongs to a translation, never to font metrics.
 * The caller additionally qualifies the complete glyph hull under its world. */
export function exactTextTranslation(x:AffineRational,y:AffineRational,maxCoordinateEmu:number,budget:SourceAffineBudget):QualifiedSourceAffine {
  const result=diagonal(one,one,checked(x,budget),checked(y,budget))
  qualifySourceAffinePoint(convertSourceAffine(result,budget).qualified,0,0,maxCoordinateEmu,budget)
  return result
}

/** Apply transform, clip referenceRect, then inverse before painting text.
 * The exact pair is identity. The existing slide clip remains in force.
 * Return rational coefficients: paint conversion is independently qualified,
 * not installed as a new approximate source-of-truth matrix. Integration must
 * also qualify the restored pair on every glyph/placeholder hull, including
 * ink outside the reference rectangle, under the final world transform. */
export function exactTextClipPlacement(area:ExactTextArea,reference:Readonly<{cx:number;cy:number}>,maxCoordinateEmu:number,budget:SourceAffineBudget):{
  rect:{x:0;y:0;cx:number;cy:number};transform:QualifiedSourceAffine;inverse:QualifiedSourceAffine
} {
  const x=checked(area.x,budget),y=checked(area.y,budget),width=checked(area.cx,budget),height=checked(area.cy,budget)
  if(width.numerator<=0n||height.numerator<=0n||!Number.isSafeInteger(reference.cx)||!Number.isSafeInteger(reference.cy)||reference.cx<=0||reference.cy<=0)throw new RangeError('Exact text clip dimensions must be positive safe references')
  budget.charge(64)
  const sx=sourceAffineRational(width.numerator,width.denominator*BigInt(reference.cx))
  const sy=sourceAffineRational(height.numerator,height.denominator*BigInt(reference.cy))
  const transform=diagonal(sx,sy,x,y)
  const inverse=diagonal(sourceAffineRational(sx.denominator,sx.numerator),sourceAffineRational(sy.denominator,sy.numerator),
    sourceAffineRational(-x.numerator*sx.denominator,x.denominator*sx.numerator),
    sourceAffineRational(-y.numerator*sy.denominator,y.denominator*sy.numerator))
  const painted=convertSourceAffine(transform,budget).qualified,undo=convertSourceAffine(inverse,budget).qualified
  const restored=composeSourceAffines(painted,undo,budget)
  for(const [px,py] of [[0,0],[reference.cx,0],[reference.cx,reference.cy],[0,reference.cy]] as const){
    qualifySourceAffinePoint(painted,px,py,maxCoordinateEmu,budget)
    qualifySourceAffinePoint(restored,px,py,maxCoordinateEmu,budget)
  }
  return {rect:{x:0,y:0,cx:reference.cx,cy:reference.cy},transform,inverse}
}
