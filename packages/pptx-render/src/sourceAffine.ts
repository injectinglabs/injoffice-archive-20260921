/** Preview affine arithmetic; integration into the transport is a separate step.
 * Rational algebra is exact. The noncardinal libm allowance is an explicit
 * deterministic qualification policy, not a formal proof of Math.sin/cos. */
export interface AffineRational { readonly numerator: bigint; readonly denominator: bigint }
type Six<T> = readonly [T,T,T,T,T,T]
export interface QualifiedSourceAffine { values: Six<AffineRational>; errors: Six<AffineRational>; depth: number }
export interface SourceAffineFrame { x:number; y:number; cx:number; cy:number; rotation?:number; flipH?:boolean; flipV?:boolean }
export const SOURCE_AFFINE_LIMITS = Object.freeze({ rationalBits:512, maxDepth:64, maxOperations:1_000_000, maxPayloadBytes:16*1024*1024, maxErrorEmu:0.125 })
/** One instance belongs to one compile request, shared by every precise node. */
export class SourceAffineBudget {
 private operations=0
 private bytes=0
 charge(operations:number,bytes=0):void {
  if(!Number.isSafeInteger(operations)||operations<0||!Number.isSafeInteger(bytes)||bytes<0)throw new RangeError('Invalid affine budget charge')
  if(operations>SOURCE_AFFINE_LIMITS.maxOperations-this.operations||bytes>SOURCE_AFFINE_LIMITS.maxPayloadBytes-this.bytes)throw new RangeError('Aggregate affine operation or payload budget exceeded')
  this.operations+=operations;this.bytes+=bytes
 }
}
const zero={numerator:0n,denominator:1n},one={numerator:1n,denominator:1n}
const abs=(n:bigint)=>n<0n?-n:n
function rational(n:bigint,d=1n):AffineRational {
 if(d===0n)throw new RangeError('Affine division by zero')
 if(d<0n){n=-n;d=-d}
 let a=abs(n),b=d;while(b){const r=a%b;a=b;b=r}n/=a;d/=a
 if(abs(n).toString(2).length>512||d.toString(2).length>512)throw new RangeError('Affine rational precision budget exceeded')
 return {numerator:n,denominator:d}
}
const add=(a:AffineRational,b:AffineRational)=>rational(a.numerator*b.denominator+b.numerator*a.denominator,a.denominator*b.denominator)
const neg=(a:AffineRational)=>rational(-a.numerator,a.denominator)
const sub=(a:AffineRational,b:AffineRational)=>add(a,neg(b))
const mul=(a:AffineRational,b:AffineRational)=>rational(a.numerator*b.numerator,a.denominator*b.denominator)
const absolute=(a:AffineRational)=>rational(abs(a.numerator),a.denominator)
const integer=(n:number)=>{if(!Number.isSafeInteger(n))throw new RangeError('Affine input requires safe integers');return rational(BigInt(n))}
function fromFloat(value:number):AffineRational {
 if(!Number.isFinite(value))throw new RangeError('Nonfinite affine coefficient')
 if(value===0)return zero
 const bytes=new DataView(new ArrayBuffer(8));bytes.setFloat64(0,value)
 const bits=bytes.getBigUint64(0),exponent=Number(bits>>52n&2047n),fraction=bits&((1n<<52n)-1n)
 let numerator=exponent===0?fraction:fraction+(1n<<52n),power=(exponent===0?-1022:exponent-1023)-52
 if(bits>>63n)numerator=-numerator
 return power>=0?rational(numerator<<BigInt(power)):rational(numerator,1n<<BigInt(-power))
}
type Term={value:AffineRational;error:AffineRational}
const term=(value:AffineRational,error=zero):Term=>({value,error})
const sum=(a:Term,b:Term):Term=>term(add(a.value,b.value),add(a.error,b.error))
const product=(a:Term,b:Term):Term=>term(mul(a.value,b.value),add(add(mul(absolute(a.value),b.error),mul(absolute(b.value),a.error)),mul(a.error,b.error)))
function matrix(terms:Six<Term>,depth=1):QualifiedSourceAffine{return {values:terms.map(t=>rational(t.value.numerator,t.value.denominator)) as unknown as Six<AffineRational>,errors:terms.map(t=>rational(t.error.numerator,t.error.denominator)) as unknown as Six<AffineRational>,depth}}
export function identitySourceAffine():QualifiedSourceAffine{return matrix([term(one),term(zero),term(zero),term(one),term(zero),term(zero)],0)}

/** Annex L.4.7.6: U^-1 R F U Tst. Flips precede rotation. */
export function sourceAffine(frame:SourceAffineFrame,child?:Pick<SourceAffineFrame,'x'|'y'|'cx'|'cy'>,budget=new SourceAffineBudget()):QualifiedSourceAffine {
 budget.charge(256)
 for(const value of [frame.x,frame.y,frame.cx,frame.cy,...(child?[child.x,child.y,child.cx,child.cy]:[])])integer(value)
 if(frame.cx<=0||frame.cy<=0||child&&(child.cx<=0||child.cy<=0))throw new RangeError('Affine extents must be positive')
 const angle=frame.rotation??0
 if(!Number.isInteger(angle)||angle<0||angle>=21600000)throw new RangeError('Affine rotation must be canonical DrawingML angle')
 if(frame.flipH!==undefined&&typeof frame.flipH!=='boolean'||frame.flipV!==undefined&&typeof frame.flipV!=='boolean')throw new TypeError('Affine flips must be boolean')
 const cardinal=angle%5400000===0,q=angle/5400000
 const radians=angle*Math.PI/10800000
 const sin=cardinal?[0,1,0,-1][q]!:Math.sin(radians),cos=cardinal?[1,0,-1,0][q]!:Math.cos(radians)
 const allowance=cardinal?zero:fromFloat(64*Number.EPSILON*(Math.abs(radians)+1))
 const c=term(fromFloat(cos),allowance),s=term(fromFloat(sin),allowance)
 const fx=integer(frame.flipH?-1:1),fy=integer(frame.flipV?-1:1)
 const sx=child?rational(BigInt(frame.cx),BigInt(child.cx)):one,sy=child?rational(BigInt(frame.cy),BigInt(child.cy)):one
 const a=product(c,term(mul(fx,sx))),b=product(s,term(mul(fx,sx))),cc=product(term(neg(s.value),s.error),term(mul(fy,sy))),d=product(c,term(mul(fy,sy)))
 const centerX=add(integer(frame.x),rational(BigInt(frame.cx),2n)),centerY=add(integer(frame.y),rational(BigInt(frame.cy),2n))
 const localX=child?add(integer(child.x),rational(BigInt(child.cx),2n)):rational(BigInt(frame.cx),2n)
 const localY=child?add(integer(child.y),rational(BigInt(child.cy),2n)):rational(BigInt(frame.cy),2n)
 const tx=sum(term(centerX),sum(product(a,term(neg(localX))),product(cc,term(neg(localY)))))
 const ty=sum(term(centerY),sum(product(b,term(neg(localX))),product(d,term(neg(localY)))))
 return matrix([a,b,cc,d,tx,ty])
}

export function composeSourceAffines(parent:QualifiedSourceAffine,local:QualifiedSourceAffine,budget=new SourceAffineBudget()):QualifiedSourceAffine {
 budget.charge(256)
 const depth=parent.depth+local.depth;if(depth>64)throw new RangeError('Affine depth budget exceeded')
 const p=parent.values.map((v,i)=>term(v,parent.errors[i]!)),l=local.values.map((v,i)=>term(v,local.errors[i]!))
 const pair=(i:number,j:number,k:number,m:number)=>sum(product(p[i]!,l[j]!),product(p[k]!,l[m]!))
 return matrix([pair(0,0,2,1),pair(1,0,3,1),pair(0,2,2,3),pair(1,2,3,3),sum(pair(0,4,2,5),p[4]!),sum(pair(1,4,3,5),p[5]!)],depth)
}

/** Annex L.4.7.4–5 is deliberately not ordinary parent×child geometry:
 * scales/flips multiply on the leaf's own axes and rotation angles add.
 * Ordinary nested matrices determine only the leaf's final center. Parents
 * are listed nearest-first. Keep conventional composition available for that
 * center mapping and for already-qualified legacy authored affine semantics. */
export function sourceHierarchyAffine(leaf:SourceAffineFrame,parents:readonly {frame:SourceAffineFrame;child:Pick<SourceAffineFrame,'x'|'y'|'cx'|'cy'>}[],budget=new SourceAffineBudget()):QualifiedSourceAffine {
 if(parents.length>=64)throw new RangeError('Affine depth budget exceeded')
 let centerMap=sourceAffine(leaf,undefined,budget),sx=one,sy=one,rotation=leaf.rotation??0,flipH=leaf.flipH??false,flipV=leaf.flipV??false
 for(const parent of parents){
  const local=sourceAffine(parent.frame,parent.child,budget)
  centerMap=composeSourceAffines(local,centerMap,budget)
  sx=mul(sx,rational(BigInt(parent.frame.cx),BigInt(parent.child.cx)))
  sy=mul(sy,rational(BigInt(parent.frame.cy),BigInt(parent.child.cy)))
  rotation=(rotation+(parent.frame.rotation??0))%21600000
  flipH=flipH!==(parent.frame.flipH??false);flipV=flipV!==(parent.frame.flipV??false)
 }
 budget.charge(256)
 const oriented=sourceAffine({x:0,y:0,cx:leaf.cx,cy:leaf.cy,rotation,flipH,flipV},undefined,budget)
 const coeff=oriented.values.slice(0,4).map((value,i)=>product(term(value,oriented.errors[i]!),term(i<2?sx:sy)))
 const halfX=term(rational(BigInt(leaf.cx),2n)),halfY=term(rational(BigInt(leaf.cy),2n))
 const center=(row:number)=>sum(sum(product(term(centerMap.values[row]!,centerMap.errors[row]!),halfX),product(term(centerMap.values[row+2]!,centerMap.errors[row+2]!),halfY)),term(centerMap.values[row+4]!,centerMap.errors[row+4]!))
 const subtractTerm=(a:Term,b:Term)=>sum(a,term(neg(b.value),b.error))
 const tx=subtractTerm(center(0),sum(product(coeff[0]!,halfX),product(coeff[2]!,halfY)))
 const ty=subtractTerm(center(1),sum(product(coeff[1]!,halfX),product(coeff[3]!,halfY)))
 return matrix([coeff[0]!,coeff[1]!,coeff[2]!,coeff[3]!,tx,ty],parents.length+1)
}

/** Convert for paint and retain the exact error of that binary64 conversion. */
export function convertSourceAffine(affine:QualifiedSourceAffine,budget=new SourceAffineBudget()):{matrix:Six<number>;qualified:QualifiedSourceAffine} {
 budget.charge(128)
 const numbers=affine.values.map(v=>Number(v.numerator)/Number(v.denominator)) as unknown as Six<number>
 const rounded=numbers.map(fromFloat) as unknown as Six<AffineRational>
 return {matrix:numbers,qualified:matrix(rounded.map((v,i)=>term(v,add(affine.errors[i]!,absolute(sub(v,affine.values[i]!))))) as unknown as Six<Term>,affine.depth)}
}

export function qualifySourceAffinePoint(affine:QualifiedSourceAffine,x:number,y:number,maxCoordinateEmu:number,budget=new SourceAffineBudget()):void {
 budget.charge(128)
 if(!Number.isSafeInteger(maxCoordinateEmu)||maxCoordinateEmu<1||maxCoordinateEmu>281474976710655)throw new RangeError('Invalid affine coordinate budget')
 const coords=[integer(x),integer(y)]
 for(const row of [0,1]){
  const value=add(add(mul(affine.values[row]!,coords[0]!),mul(affine.values[row+2]!,coords[1]!)),affine.values[row+4]!)
  const error=add(add(mul(affine.errors[row]!,absolute(coords[0]!)),mul(affine.errors[row+2]!,absolute(coords[1]!))),affine.errors[row+4]!)
  // Reserve paint arithmetic/composition allowance in addition to exact input
  // and conversion error. This is a disclosed bounded preview policy.
  const magnitude=add(add(absolute(mul(affine.values[row]!,coords[0]!)),absolute(mul(affine.values[row+2]!,coords[1]!))),absolute(affine.values[row+4]!))
  const total=add(error,mul(magnitude,fromFloat(32*Number.EPSILON*Math.max(1,affine.depth))))
  if(total.numerator*8n>total.denominator)throw new RangeError('Affine accumulated uncertainty exceeds 0.125 EMU')
  const bound=add(absolute(value),total)
  if(bound.numerator>BigInt(maxCoordinateEmu)*bound.denominator)throw new RangeError('Affine world coordinate budget exceeded')
 }
}

export interface SourceAffineTransport {profile:'rational-affine-v1';values:Six<readonly [string,string]>;errors:Six<readonly [string,string]>;depth:number}
export function encodeSourceAffine(affine:QualifiedSourceAffine,budget=new SourceAffineBudget()):SourceAffineTransport {
 const pairs=(values:Six<AffineRational>)=>values.map(v=>{if(abs(v.numerator)>=(1n<<512n)||v.denominator<=0n||v.denominator>=(1n<<512n))throw new RangeError('Affine rational precision budget exceeded');return [v.numerator.toString(),v.denominator.toString()] as const}) as unknown as Six<readonly [string,string]>
 const transport:SourceAffineTransport={profile:'rational-affine-v1',values:pairs(affine.values),errors:pairs(affine.errors),depth:affine.depth}
 // The decoder validates canonical reduction/sign/bit budgets for local values
 // too, so encoding cannot smuggle an invalid caller-constructed affine.
 decodeSourceAffine(transport,budget)
 return transport
}
export function decodeSourceAffine(input:unknown,budget=new SourceAffineBudget()):QualifiedSourceAffine {
 if(!input||typeof input!=='object'||Array.isArray(input))throw new TypeError('Invalid affine transport')
 const fields=Object.getOwnPropertyDescriptors(input),keys=Reflect.ownKeys(fields)
 if(keys.length!==4||!['profile','values','errors','depth'].every(k=>k in fields&&'value' in fields[k]!))throw new TypeError('Invalid affine transport fields')
 const profile=fields.profile!.value,depth=fields.depth!.value
 if(profile!=='rational-affine-v1'||!Number.isInteger(depth)||depth<0||depth>64)throw new RangeError('Invalid affine transport profile or depth')
 let bytes=128
 const tuple=(raw:unknown,length:number):unknown[]=>{
  if(!Array.isArray(raw)||raw.length!==length||Reflect.ownKeys(raw).length!==length+1)throw new TypeError('Invalid affine tuple')
  return Array.from({length},(_,index)=>{const descriptor=Object.getOwnPropertyDescriptor(raw,String(index));if(!descriptor||!('value' in descriptor))throw new TypeError('Affine tuple holes or accessors are unsupported');return descriptor.value})
 }
 const parse=(raw:unknown,error:boolean):Six<AffineRational>=>{
  return tuple(raw,6).map(rawPair=>{
   const pair=tuple(rawPair,2)
   if(pair.some(v=>typeof v!=='string'||v.length>156))throw new TypeError('Invalid affine rational encoding')
   const [n,d]=pair as [string,string]
   if(!/^(?:0|-?[1-9][0-9]*)$/.test(n)||!/^[1-9][0-9]*$/.test(d))throw new TypeError('Noncanonical affine rational')
   bytes+=n.length+d.length+16
   const value=rational(BigInt(n),BigInt(d))
   if(value.numerator.toString()!==n||value.denominator.toString()!==d||error&&value.numerator<0n)throw new TypeError('Noncanonical affine rational or negative error')
   return value
  }) as unknown as Six<AffineRational>
 }
 const result={values:parse(fields.values!.value,false),errors:parse(fields.errors!.value,true),depth}
 budget.charge(128,bytes)
 return result
}
