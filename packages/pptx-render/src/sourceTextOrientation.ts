import {composeSourceAffines, identitySourceAffine, sourceAffine, SourceAffineBudget, SOURCE_AFFINE_LIMITS, type AffineRational, type QualifiedSourceAffine, type SourceAffineFrame} from './sourceAffine.js'

type SourceParent = {frame:SourceAffineFrame;child:Pick<SourceAffineFrame,'x'|'y'|'cx'|'cy'>}
const zero:AffineRational={numerator:0n,denominator:1n}
const one:AffineRational={numerator:1n,denominator:1n}
function fraction(n:bigint,d:bigint):AffineRational {
  if(d<=0n)throw new RangeError('Text orientation requires a positive rational denominator')
  let a=n<0n?-n:n,b=d
  while(b){const remainder=a%b;a=b;b=remainder}
  n/=a;d/=a
  if((n<0n?-n:n).toString(2).length>SOURCE_AFFINE_LIMITS.rationalBits||d.toString(2).length>SOURCE_AFFINE_LIMITS.rationalBits)throw new RangeError('Text orientation rational precision budget exceeded')
  return {numerator:n,denominator:d}
}
function centeredScale(cx:number,cy:number,sx:AffineRational,sy:AffineRational):QualifiedSourceAffine {
  return {values:[sx,zero,zero,sy,
    fraction(BigInt(cx)*(sx.denominator-sx.numerator),2n*sx.denominator),
    fraction(BigInt(cy)*(sy.denominator-sy.numerator),2n*sy.denominator)],errors:[zero,zero,zero,zero,zero,zero],depth:1}
}

/** Additional text-local transform before vertical writing-mode layout.
 * Body rotation acts on the group-scaled text anchor (the pinned POI preview
 * policy), rather than introducing a shear by rotating below unequal scales.
 * Upright requires a separate text-area layout decision and is not handled here.
 * The caller must share the compile budget and qualify the resulting glyph and
 * clip hulls after composing this result with the source hierarchy transform. */
export function sourceBodyRotation(
  frame:SourceAffineFrame, parents:readonly SourceParent[], rotation:number,
  budget:SourceAffineBudget,
):QualifiedSourceAffine {
  if(!Number.isInteger(rotation)||Object.is(rotation,-0)||rotation< -2147483648||rotation>2147483647)throw new RangeError('Text rotation requires a signed DrawingML angle')
  if(parents.length>=SOURCE_AFFINE_LIMITS.maxDepth)throw new RangeError('Text orientation depth budget exceeded')
  // Validate the complete frame inputs even when the resulting angle is zero.
  sourceAffine(frame,undefined,budget)
  let sx=one,sy=one,reflected=(frame.flipH??false)!==(frame.flipV??false)
  for(const parent of parents){
    sourceAffine(parent.frame,parent.child,budget)
    budget.charge(16)
    sx=fraction(sx.numerator*BigInt(parent.frame.cx),sx.denominator*BigInt(parent.child.cx))
    sy=fraction(sy.numerator*BigInt(parent.frame.cy),sy.denominator*BigInt(parent.child.cy))
    reflected=reflected!==((parent.frame.flipH??false)!==(parent.frame.flipV??false))
  }
  const localFrame={x:0,y:0,cx:frame.cx,cy:frame.cy}
  const counter=reflected?sourceAffine({...localFrame,flipH:true},undefined,budget):identitySourceAffine()
  const angle=(rotation%21600000+21600000)%21600000
  if(angle===0)return counter
  const turn=sourceAffine({...localFrame,rotation:angle},undefined,budget)
  const scale=centeredScale(frame.cx,frame.cy,sx,sy)
  const inverseScale=centeredScale(frame.cx,frame.cy,{numerator:sx.denominator,denominator:sx.numerator},{numerator:sy.denominator,denominator:sy.numerator})
  // C · S^-1 · Rbody · S: the node already supplies Rshape · F · S.
  return composeSourceAffines(counter,composeSourceAffines(inverseScale,composeSourceAffines(turn,scale,budget),budget),budget)
}

export interface ExactTextArea {
  readonly x:AffineRational;readonly y:AffineRational
  readonly cx:AffineRational;readonly cy:AffineRational
}
export const UPRIGHT_TEXT_AREA_POLICY='source-upright-physical-quadrants-v1'

/** Exact proposal boundary: do not convert these dimensions to an integer
 * rectangle or rescale glyphs to fit it. Layout must compare rational available
 * widths and retain rational anchor offsets before final paint qualification.
 * S (and therefore supplied glyph metrics) is unchanged. Only the available
 * text area exchanges physical width/height in the near-vertical quadrants. */
export function sourceUprightTextArea(
  frame:SourceAffineFrame,parents:readonly SourceParent[],
  content:Readonly<{x:number;y:number;cx:number;cy:number}>,budget:SourceAffineBudget,
):{area:ExactTextArea;orientation:QualifiedSourceAffine;swapped:boolean;policy:typeof UPRIGHT_TEXT_AREA_POLICY} {
  if(parents.length>=SOURCE_AFFINE_LIMITS.maxDepth)throw new RangeError('Text orientation depth budget exceeded')
  sourceAffine(frame,undefined,budget)
  for(const value of [content.x,content.y,content.cx,content.cy])if(!Number.isSafeInteger(value))throw new RangeError('Text content area requires safe integers')
  if(content.cx<=0||content.cy<=0)throw new RangeError('Text content area must be positive')
  let sx=one,sy=one,angle=frame.rotation??0,flipH=frame.flipH??false,flipV=frame.flipV??false
  for(const parent of parents){
    sourceAffine(parent.frame,parent.child,budget);budget.charge(16)
    sx=fraction(sx.numerator*BigInt(parent.frame.cx),sx.denominator*BigInt(parent.child.cx))
    sy=fraction(sy.numerator*BigInt(parent.frame.cy),sy.denominator*BigInt(parent.child.cy))
    angle=(angle+(parent.frame.rotation??0))%21600000
    flipH=flipH!==(parent.frame.flipH??false);flipV=flipV!==(parent.frame.flipV??false)
  }
  budget.charge(64)
  const swapped=angle>=2700000&&angle<8100000||angle>=13500000&&angle<18900000
  // S·(new width,new height) = (old physical height,old physical width).
  const cx=swapped?fraction(BigInt(content.cy)*sy.numerator*sx.denominator,sy.denominator*sx.numerator):fraction(BigInt(content.cx),1n)
  const cy=swapped?fraction(BigInt(content.cx)*sx.numerator*sy.denominator,sx.denominator*sy.numerator):fraction(BigInt(content.cy),1n)
  const x=fraction((2n*BigInt(content.x)+BigInt(content.cx))*cx.denominator-cx.numerator,2n*cx.denominator)
  const y=fraction((2n*BigInt(content.y)+BigInt(content.cy))*cy.denominator-cy.numerator,2n*cy.denominator)
  const localFrame={x:0,y:0,cx:frame.cx,cy:frame.cy}
  const undoTurn=sourceAffine({...localFrame,rotation:(21600000-angle)%21600000},undefined,budget)
  const flips=sourceAffine({...localFrame,flipH,flipV},undefined,budget)
  const scale=centeredScale(frame.cx,frame.cy,sx,sy)
  const inverseScale=centeredScale(frame.cx,frame.cy,{numerator:sx.denominator,denominator:sx.numerator},{numerator:sy.denominator,denominator:sy.numerator})
  // R·F·S already belongs to the leaf. Local S^-1·F·R^-1·S leaves
  // the original positive glyph scale S, with neither mirroring nor rotation.
  const orientation=composeSourceAffines(inverseScale,composeSourceAffines(flips,composeSourceAffines(undoTurn,scale,budget),budget),budget)
  return {area:{x,y,cx,cy},orientation,swapped,policy:UPRIGHT_TEXT_AREA_POLICY}
}
