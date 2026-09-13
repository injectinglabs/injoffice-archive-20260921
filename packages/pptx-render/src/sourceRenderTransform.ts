import {convertSourceAffine,decodeSourceAffine,encodeSourceAffine,SourceAffineBudget,type QualifiedSourceAffine} from './sourceAffine.js'
import type {RenderTransform} from './types.js'

/** Preserve the established integer-PPM transport when all coefficients and
 * translations are exactly representable. Otherwise identity PPM fields are
 * required and the explicit rational variant is authoritative. */
export function sourceRenderTransform(affine:QualifiedSourceAffine,budget:SourceAffineBudget):RenderTransform {
 const integers=affine.values.map((v,i)=>{const n=v.numerator*(i<4?1000000n:1n);return n%v.denominator===0n?n/v.denominator:undefined})
 if(affine.errors.every(v=>v.numerator===0n)&&integers.every(v=>v!==undefined&&v>=-9007199254740991n&&v<=9007199254740991n)) {
  const [aPpm,bPpm,cPpm,dPpm,txEmu,tyEmu]=integers.map(Number)
  return {aPpm:aPpm!,bPpm:bPpm!,cPpm:cPpm!,dPpm:dPpm!,txEmu:txEmu!,tyEmu:tyEmu!}
 }
 return {aPpm:1000000,bPpm:0,cPpm:0,dPpm:1000000,txEmu:0,tyEmu:0,sourceAffine:encodeSourceAffine(affine,budget)}
}

/** Consumer boundary: decode the rational variant before binary64 conversion.
 * Callers share one budget across a complete render/paint request. */
export function renderTransformMatrix(transform:RenderTransform,budget:SourceAffineBudget):readonly [number,number,number,number,number,number] {
 if(transform.sourceAffine!==undefined){
  if(transform.aPpm!==1000000||transform.bPpm!==0||transform.cPpm!==0||transform.dPpm!==1000000||transform.txEmu!==0||transform.tyEmu!==0)throw new TypeError('Rational transform cannot also carry a PPM operation')
  return convertSourceAffine(decodeSourceAffine(transform.sourceAffine,budget),budget).matrix
 }
 const values=[transform.aPpm,transform.bPpm,transform.cPpm,transform.dPpm,transform.txEmu,transform.tyEmu]
 if(values.some(v=>!Number.isSafeInteger(v)))throw new TypeError('Invalid integer PPM transform')
 return [transform.aPpm/1e6,transform.bPpm/1e6,transform.cPpm/1e6,transform.dPpm/1e6,transform.txEmu,transform.tyEmu]
}
