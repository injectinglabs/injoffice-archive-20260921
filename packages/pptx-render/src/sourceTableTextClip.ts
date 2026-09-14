import {convertSourceAffine,qualifySourceAffinePoint,sourceAffineRational,addSourceAffineRationals,SourceAffineBudget,type AffineRational,type QualifiedSourceAffine} from './sourceAffine.js'

type Rect=Readonly<{x:number;y:number;cx:number;cy:number}>
const floor=(n:bigint,d:bigint)=>n/d-(n<0n&&n%d!==0n?1n:0n)
const multiple=(a:AffineRational,n:number)=>sourceAffineRational(a.numerator*BigInt(n),a.denominator)

/** A finite cell-local horizontal strip. The unchanged slide clip supplies the
 * outer boundary; y covers all ink, preserving vertical overflow after any
 * ancestor transform. Callers include measured glyph/control hulls and separate
 * refusal-placeholder hulls with their actual local paint transforms. Background
 * and border paint must remain outside this text-only clip. The caller also
 * qualifies the returned clip corners under the final ancestor/world matrix.
 * No source or table admission guard is relaxed by this helper. */
export function sourceTableTextClip(
  cellWidth:number,spans:readonly {bounds:Rect;transform:QualifiedSourceAffine}[],
  maxCoordinateEmu:number,budget:SourceAffineBudget,
):Rect {
  if(!Number.isSafeInteger(cellWidth)||cellWidth<=0||cellWidth>maxCoordinateEmu||spans.length===0)throw new RangeError('Invalid table text clip input')
  let minimum:bigint|undefined,maximum:bigint|undefined
  for(const span of spans){
    budget.charge(16)
    const {x,y,cx,cy}=span.bounds
    if([x,y,cx,cy,x+cx,y+cy].some(value=>!Number.isSafeInteger(value))||cx<0||cy<0)throw new RangeError('Invalid table text control hull')
    const converted=convertSourceAffine(span.transform,budget).qualified
    for(const [px,py] of [[x,y],[x+cx,y],[x+cx,y+cy],[x,y+cy]] as const){
      qualifySourceAffinePoint(converted,px,py,maxCoordinateEmu,budget)
      const value=addSourceAffineRationals(addSourceAffineRationals(multiple(converted.values[1],px),multiple(converted.values[3],py)),converted.values[5])
      // The qualifier above bounds all accumulated/conversion/paint error by
      // 1/8 EMU. Enclose that whole allowance before outward integer rounding.
      const low=floor(value.numerator*8n-value.denominator,value.denominator*8n)
      const high=-floor(-(value.numerator*8n+value.denominator),value.denominator*8n)
      minimum=minimum===undefined||low<minimum?low:minimum
      maximum=maximum===undefined||high>maximum?high:maximum
    }
  }
  const lower=Number(minimum!),upper=Number(maximum!),height=upper-lower
  if(!Number.isSafeInteger(lower)||!Number.isSafeInteger(upper)||!Number.isSafeInteger(height)||height<=0||Math.abs(lower)>maxCoordinateEmu||Math.abs(upper)>maxCoordinateEmu||height>maxCoordinateEmu)throw new RangeError('Table text clip exceeds local coordinate budget')
  return {x:0,y:lower,cx:cellWidth,cy:height}
}
