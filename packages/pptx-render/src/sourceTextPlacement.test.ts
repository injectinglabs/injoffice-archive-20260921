import {describe,expect,it} from 'vitest'
import {composeSourceAffines,convertSourceAffine,identitySourceAffine,SourceAffineBudget} from './sourceAffine.js'
import {exactTextClipPlacement,exactTextTranslation} from './sourceTextPlacement.js'
const r=(numerator:bigint,denominator=1n)=>({numerator,denominator})
const area={x:r(230n,3n),y:r(-5n,2n),cx:r(200n,3n),cy:r(225n)}
describe('exact text placement preparation',()=>{
  it('retains fractional alignment origins as translations without changing glyph axes',()=>{
    const result=exactTextTranslation(r(1n,3n),r(-5n,2n),10000,new SourceAffineBudget())
    expect(result.values).toEqual([r(1n),r(0n),r(0n),r(1n),r(1n,3n),r(-5n,2n)])
  })
  it('maps integer reference corners to the exact fractional clip and restores identity',()=>{
    const budget=new SourceAffineBudget()
    const result=exactTextClipPlacement(area,{cx:150,cy:100},10000,budget)
    expect(result.transform.values).toEqual([r(4n,9n),r(0n),r(0n),r(9n,4n),r(230n,3n),r(-5n,2n)])
    const restored=composeSourceAffines(result.transform,result.inverse,budget)
    expect(restored.values).toEqual(identitySourceAffine().values)
    expect(restored.errors.every(value=>value.numerator===0n)).toBe(true)
    const [a,b,c,d,e,f]=convertSourceAffine(result.transform,budget).matrix
    expect(a*150+c*100+e).toBeCloseTo(430/3,12)
    expect(b*150+d*100+f).toBe(445/2)
  })
  it('rejects invalid, collapsed, ill-conditioned or out-of-budget placements',()=>{
    for(const bad of [{...area,cx:r(0n)},{...area,cy:r(-1n)},{...area,x:r(1n,0n)}])expect(()=>exactTextClipPlacement(bad,{cx:150,cy:100},10000,new SourceAffineBudget())).toThrow()
    expect(()=>exactTextClipPlacement(area,{cx:0,cy:1},10000,new SourceAffineBudget())).toThrow()
    expect(()=>exactTextTranslation(r(10001n),r(0n),10000,new SourceAffineBudget())).toThrow(/coordinate/)
    expect(()=>exactTextClipPlacement({...area,x:r(100000000000000n)},{cx:150,cy:100},281474976710655,new SourceAffineBudget())).toThrow(/uncertainty/)
    const budget=new SourceAffineBudget();budget.charge(1_000_000)
    expect(()=>exactTextClipPlacement(area,{cx:150,cy:100},10000,budget)).toThrow(/budget/)
  })
})
