import {expect,it} from 'vitest'
import {chartRational as r} from './chartRational.js'
import {chartRoundedSquareRoot as root} from './chartBoundedSqrt.js'
it('rounds exact squares and half thresholds without binary cancellation',()=>{
 expect(root(r(0n))).toBe(0);expect(root(r(4n))).toBe(2)
 expect(root(r(25n,4n))).toBe(3)
 expect(root(r(249999999999999999999999999999999n,40000000000000000000000000000000n))).toBe(2)
 expect(root(r(250000000000000000000000000000001n,40000000000000000000000000000000n))).toBe(3)
 const big=281474976710655n;expect(root(r(big*big))).toBe(Number(big))
 expect(root(r(1n,10n**500n))).toBe(0)
})
it('matches independently squared rounding intervals over bounded rational inputs',()=>{
 for(let n=0;n<500;n++)for(let d=1;d<17;d++){
  const result=BigInt(root(r(BigInt(n),BigInt(d))))
  expect(4n*BigInt(n)<(2n*result+1n)**2n*BigInt(d)).toBe(true)
  if(result>0n)expect(4n*BigInt(n)>=(2n*result-1n)**2n*BigInt(d)).toBe(true)
 }
})
it('refuses invalid or over-budget roots',()=>{
 expect(()=>root(r(-1n))).toThrow();expect(()=>root({numerator:1n,denominator:0n})).toThrow()
 expect(()=>root({numerator:1n<<4096n,denominator:1n})).toThrow()
 expect(()=>root(r((BigInt(Number.MAX_SAFE_INTEGER)+1n)**2n))).toThrow()
})
