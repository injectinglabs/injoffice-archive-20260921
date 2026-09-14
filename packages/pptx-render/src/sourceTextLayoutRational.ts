import {sourceAffineRational,addSourceAffineRationals,subtractSourceAffineRationals,SOURCE_AFFINE_LIMITS,SourceAffineBudget,type AffineRational} from './sourceAffine.js'

function checked(value:AffineRational,budget:SourceAffineBudget):AffineRational {
  budget.charge(16)
  if(!value||typeof value.numerator!=='bigint'||typeof value.denominator!=='bigint'||value.denominator<=0n)throw new RangeError('Invalid exact text length')
  if((value.numerator<0n?-value.numerator:value.numerator).toString(2).length>SOURCE_AFFINE_LIMITS.rationalBits||value.denominator.toString(2).length>SOURCE_AFFINE_LIMITS.rationalBits)throw new RangeError('Exact text precision budget exceeded')
  return sourceAffineRational(value.numerator,value.denominator)
}
function integer(value:number):AffineRational {
  if(!Number.isSafeInteger(value)||Object.is(value,-0))throw new RangeError('Exact text advance or margin requires a safe integer')
  return sourceAffineRational(BigInt(value))
}

/** Subtract native integer paragraph margins without changing the rational
 * wrapping threshold. A width of 200/3 EMU must not become 67 before fitting. */
export function exactTextLineWidth(width:AffineRational,marginEmu:number,budget:SourceAffineBudget):AffineRational {
  const result=subtractSourceAffineRationals(checked(width,budget),integer(marginEmu))
  if(result.numerator<=0n)throw new RangeError('Paragraph margins leave no positive text line width')
  return result
}

export function exactTextAdvanceFits(advanceEmu:number,width:AffineRational,budget:SourceAffineBudget):boolean {
  const available=checked(width,budget),advance=integer(advanceEmu)
  if(available.numerator<=0n)throw new RangeError('Exact text line width must be positive')
  return advance.numerator*available.denominator<=available.numerator
}

/** Retain fractional alignment offsets until the paint-coordinate boundary.
 * Negative offsets for overflowing lines remain meaningful; no clamping. */
export function exactTextAlignmentOffset(align:'left'|'center'|'right',width:AffineRational,advanceEmu:number,budget:SourceAffineBudget):AffineRational {
  const available=checked(width,budget)
  if(available.numerator<=0n)throw new RangeError('Exact text line width must be positive')
  const remainder=subtractSourceAffineRationals(available,integer(advanceEmu))
  if(align==='left')return sourceAffineRational(0n)
  if(align==='center')return sourceAffineRational(remainder.numerator,remainder.denominator*2n)
  if(align==='right')return remainder
  throw new RangeError('Unknown exact text alignment')
}

export function exactTextAnchorOffset(anchor:'top'|'center'|'bottom',height:AffineRational,usedHeightEmu:number,budget:SourceAffineBudget):AffineRational {
  const available=checked(height,budget)
  if(available.numerator<=0n)throw new RangeError('Exact text area height must be positive')
  const remainder=subtractSourceAffineRationals(available,integer(usedHeightEmu))
  if(anchor==='top')return sourceAffineRational(0n)
  if(anchor==='center')return sourceAffineRational(remainder.numerator,remainder.denominator*2n)
  if(anchor==='bottom')return remainder
  throw new RangeError('Unknown exact text anchor')
}

/** Keep the area's fractional origin separate from font advances and metrics. */
export function exactTextOffsetSum(origin:AffineRational,offset:AffineRational,budget:SourceAffineBudget):AffineRational {
  return addSourceAffineRationals(checked(origin,budget),checked(offset,budget))
}
