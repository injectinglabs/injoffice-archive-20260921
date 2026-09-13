import {expect,it} from 'vitest'
import {chartRational as r,chartRationalAdd as add,chartRationalSubtract as sub,chartRationalMultiply as mul,chartRationalDivide as div,chartRationalCompare as cmp,chartRationalDecimal as decimal,chartRationalCoordinate as coordinate} from './chartRational.js'
it('reduces signed exact fractions and preserves decimal cancellation',()=>{
 expect(r(-6n,-8n)).toEqual(r(3n,4n));expect(r(0n,-9n)).toEqual(r(0n))
 const a=decimal('1000000000000.000000000000000001'),b=decimal('1000000000000.000000000000000002')
 expect(sub(b,a)).toEqual(decimal('1e-18'));expect(cmp(a,b)).toBe(-1)
 expect(add(r(1n,3n),r(1n,6n))).toEqual(r(1n,2n));expect(mul(r(2n,3n),r(9n,4n))).toEqual(r(3n,2n));expect(div(r(2n,3n),r(4n,9n))).toEqual(r(3n,2n))
})
it('rounds only final coordinates and applies reversal before half ties',()=>{
 expect(coordinate(r(1n,2n),1001)).toBe(501);expect(coordinate(r(1n,2n),1001,true)).toBe(501)
 expect(coordinate(r(1n,3n),1000)).toBe(333);expect(coordinate(r(1n,3n),1000,true)).toBe(667)
 for(const fraction of [r(-1n),r(2n)])expect(()=>coordinate(fraction,1000)).toThrow()
 expect(()=>r(1n,0n)).toThrow();expect(()=>div(r(1n),r(0n))).toThrow();expect(()=>r(1n<<4096n)).toThrow(/bit budget/)
})
