import {describe,expect,it} from 'vitest'
import {SourceAffineBudget} from './sourceAffine.js'
import {exactTextLineWidth,exactTextAdvanceFits,exactTextAlignmentOffset,exactTextAnchorOffset,exactTextOffsetSum} from './sourceTextLayoutRational.js'

const rat=(numerator:bigint,denominator=1n)=>({numerator,denominator})
describe('exact text area layout arithmetic',()=>{
  it('compares integer font advances to fractional wrap thresholds without rounding',()=>{
    const budget=new SourceAffineBudget()
    expect(exactTextAdvanceFits(66,rat(200n,3n),budget)).toBe(true)
    expect(exactTextAdvanceFits(67,rat(200n,3n),budget)).toBe(false)
    expect(exactTextAdvanceFits(67,rat(201n,3n),budget)).toBe(true)
    // The difference is much smaller than a binary64 ULP at MAXSAFE.
    const max=9007199254740991n,den=1n<<200n
    expect(exactTextAdvanceFits(Number(max),rat(max*den-1n,den),budget)).toBe(false)
    expect(exactTextAdvanceFits(Number(max),rat(max*den,den),budget)).toBe(true)
  })
  it('retains first-line and continuation margins independently',()=>{
    const budget=new SourceAffineBudget(),width=Object.freeze(rat(200n,3n))
    expect(exactTextLineWidth(width,2,budget)).toEqual(rat(194n,3n))
    expect(exactTextLineWidth(width,-1,budget)).toEqual(rat(203n,3n))
    expect(width).toEqual(rat(200n,3n))
    expect(()=>exactTextLineWidth(width,67,budget)).toThrow(/no positive/)
  })
  it('preserves exact alignment, including an overflowing centered line',()=>{
    const budget=new SourceAffineBudget(),width=rat(200n,3n)
    expect(exactTextAlignmentOffset('left',width,66,budget)).toEqual(rat(0n))
    expect(exactTextAlignmentOffset('center',width,66,budget)).toEqual(rat(1n,3n))
    expect(exactTextAlignmentOffset('right',width,66,budget)).toEqual(rat(2n,3n))
    expect(exactTextAlignmentOffset('center',width,67,budget)).toEqual(rat(-1n,6n))
  })
  it('keeps fractional anchor offsets and asymmetric origins separate from glyph metrics',()=>{
    const budget=new SourceAffineBudget(),height=rat(451n,2n)
    expect(exactTextAnchorOffset('top',height,225,budget)).toEqual(rat(0n))
    expect(exactTextAnchorOffset('center',height,225,budget)).toEqual(rat(1n,4n))
    expect(exactTextAnchorOffset('bottom',height,225,budget)).toEqual(rat(1n,2n))
    expect(exactTextOffsetSum(rat(-5n,2n),rat(1n,4n),budget)).toEqual(rat(-9n,4n))
  })
  it('rejects malformed inputs and bounds both arithmetic precision and the shared request',()=>{
    const budget=new SourceAffineBudget()
    for(const width of [rat(1n,0n),rat(1n,-1n),rat(1n<<512n),rat(1n,1n<<512n)])expect(()=>exactTextAdvanceFits(1,width,budget)).toThrow()
    for(const value of [-0,0.5,NaN,Infinity,9007199254740992])expect(()=>exactTextAdvanceFits(value,rat(1n),budget)).toThrow()
    expect(()=>exactTextOffsetSum(rat((1n<<512n)-1n),rat(1n),budget)).toThrow(/precision/)
    const exhausted=new SourceAffineBudget();exhausted.charge(1_000_000)
    expect(()=>exactTextAdvanceFits(1,rat(1n),exhausted)).toThrow(/budget/)
  })
})
