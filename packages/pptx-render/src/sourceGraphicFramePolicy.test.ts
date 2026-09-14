import {expect,it} from 'vitest'
import {SourceAffineBudget,sourceHierarchyAffine} from './sourceAffine.js'
import {sourceRenderTransform} from './sourceRenderTransform.js'
import {projectSourceGraphicFrameLayout,sourceGraphicFrameLayout,qualifySourceGraphicFrameSpans,sourceGraphicFramePaintTransform} from './sourceGraphicFramePolicy.js'
const frame=(x=0)=>sourceHierarchyAffine({x,y:0,cx:100,cy:100},[],new SourceAffineBudget())
it('checks outside-frame control hulls under the complete world mapping',()=>{
 const budget=new SourceAffineBudget(),world=frame(800),local=frame()
 expect(()=>qualifySourceGraphicFrameSpans(world,[{bounds:{x:0,y:0,cx:100,cy:100},transform:local}],1000,budget)).not.toThrow()
 expect(()=>qualifySourceGraphicFrameSpans(world,[{bounds:{x:0,y:0,cx:201,cy:100},transform:local}],1000,budget)).toThrow()
})
it('decodes precise paint and rejects combined PPM/rational authority',()=>{
 const budget=new SourceAffineBudget(),transform=sourceRenderTransform(sourceHierarchyAffine({x:0,y:0,cx:100,cy:100,rotation:1800000},[],budget),budget)
 expect(()=>sourceGraphicFramePaintTransform(transform,budget)).not.toThrow()
 expect(()=>sourceGraphicFramePaintTransform({...transform,txEmu:1},budget)).toThrow('cannot also')
})
it('bounds sparse spans and shares the request operation cap',()=>{
 expect(()=>qualifySourceGraphicFrameSpans(frame(),Array(1),1000,new SourceAffineBudget())).toThrow('Missing')
 const budget=new SourceAffineBudget();budget.charge(999999)
 expect(()=>qualifySourceGraphicFrameSpans(frame(),[{bounds:{x:0,y:0,cx:1,cy:1},transform:frame()}],1000,budget)).toThrow('operation')
})

const value=(r:{numerator:bigint;denominator:bigint})=>Number(r.numerator)/Number(r.denominator)
const group={frame:{x:3000000,y:2200000,cx:6000000,cy:2000000},child:{x:0,y:0,cx:4000000,cy:2000000}}
it('keeps direct graphic-frame paint unrotated and unreflected',()=>{
 const result=sourceGraphicFrameLayout({x:3,y:5,cx:7,cy:9,rotation:1800000,flipH:true,flipV:true},[],new SourceAffineBudget())
 expect(result.origin.values.map(value)).toEqual([1,0,0,1,3,5])
 expect([value(result.width),value(result.height)]).toEqual([7,9])
})
it('uses exact half-open raw-angle quadrants for physical frame sizing',()=>{
 for(const [angle,swap] of [[2699999,false],[2700000,true],[8099999,true],[8100000,false],[13499999,false],[13500000,true],[18899999,true],[18900000,false]] as const){
  const result=sourceGraphicFrameLayout({x:0,y:0,cx:4000000,cy:2000000,rotation:angle},[group],new SourceAffineBudget())
  expect(result.origin.values.map(value)).toEqual([1,0,0,1,swap?4000000:3000000,swap?1700000:2200000])
  expect([value(result.width),value(result.height)]).toEqual(swap?[4000000,3000000]:[6000000,2000000])
 }
})
it('maps offset centers through rotation and reflection without transforming glyph axes',()=>{
 const leaf={x:500000,y:200000,cx:4000000,cy:2000000}
 const reflected=sourceGraphicFrameLayout(leaf,[{...group,frame:{...group.frame,flipH:true}}],new SourceAffineBudget())
 expect(reflected.origin.values.map(value)).toEqual([1,0,0,1,2250000,2400000])
 const rotated=sourceGraphicFrameLayout(leaf,[{...group,frame:{...group.frame,rotation:1800000}}],new SourceAffineBudget())
 expect(value(rotated.origin.values[4])).toBeCloseTo(3549519.052838329,7)
 expect(value(rotated.origin.values[5])).toBeCloseTo(2748205.080756888,7)
 expect(rotated.origin.values.slice(0,4).map(value)).toEqual([1,0,0,1])
 expect(rotated.origin.errors[4].numerator).toBeGreaterThan(0n)
})
it('retains fractional physical layout dimensions and origins without integer coercion',()=>{
 const result=sourceGraphicFrameLayout({x:0,y:0,cx:1,cy:1},[{frame:{x:0,y:0,cx:2,cy:5},child:{x:0,y:0,cx:3,cy:7}}],new SourceAffineBudget())
 expect(result.width).toEqual({numerator:2n,denominator:3n});expect(result.height).toEqual({numerator:5n,denominator:7n})
 expect(result.origin.values.map(value)).toEqual([1,0,0,1,0,0])
})
it('rejects sparse ancestors and preserves shared numerical budgets',()=>{
 expect(()=>sourceGraphicFrameLayout({x:0,y:0,cx:1,cy:1},Array(1),new SourceAffineBudget())).toThrow('Missing')
 expect(()=>sourceGraphicFrameLayout({x:0,y:0,cx:1,cy:1,rotation:-1},[],new SourceAffineBudget())).toThrow('canonical')
 const budget=new SourceAffineBudget();budget.charge(999999)
 expect(()=>sourceGraphicFrameLayout({x:0,y:0,cx:1,cy:1},[],budget)).toThrow('operation')
})
it('keeps nested noncommuting center mapping separate from positive physical scale products',()=>{
 const result=sourceGraphicFrameLayout({x:0,y:0,cx:2,cy:2},[
  {frame:{x:0,y:0,cx:8,cy:4,rotation:5400000},child:{x:0,y:0,cx:4,cy:4}},
  {frame:{x:0,y:0,cx:24,cy:4},child:{x:0,y:0,cx:8,cy:4}},
 ],new SourceAffineBudget())
 // (1,1) -> inner quarter-turn center (5,0) -> outer scale (15,0).
 // Physical dimensions use positive products (6,1), not rotated matrix columns.
 expect([value(result.width),value(result.height)]).toEqual([12,2])
 expect(result.origin.values.map(value)).toEqual([1,0,0,1,9,-1])
 expect(result.origin.errors.every(r=>r.numerator===0n)).toBe(true)
})

it('projects physical dimensions once with exact error and outward hull allowances',()=>{
 const result=projectSourceGraphicFrameLayout({width:{numerator:5n,denominator:2n},height:{numerator:7n,denominator:3n}},100,new SourceAffineBudget())
 expect(result).toEqual({policy:'nearest-emu-physical-layout-v1',widthEmu:3,heightEmu:2,widthError:{numerator:1n,denominator:2n},heightError:{numerator:1n,denominator:3n},hullOutsetXEmu:1,hullOutsetYEmu:1})
 const exact=projectSourceGraphicFrameLayout({width:{numerator:6n,denominator:2n},height:{numerator:2n,denominator:1n}},100,new SourceAffineBudget())
 expect(exact.hullOutsetXEmu).toBe(0);expect(exact.hullOutsetYEmu).toBe(0)
})
it('refuses collapsed or over-budget projected dimensions without inventing a minimum',()=>{
 const dimension=(n:bigint,d=1n)=>({width:{numerator:n,denominator:d},height:{numerator:1n,denominator:1n}})
 for(const input of [dimension(1n,3n),dimension(201n,2n),dimension(-1n),dimension(1n,-2n)])expect(()=>projectSourceGraphicFrameLayout(input,100,new SourceAffineBudget())).toThrow()
 expect(projectSourceGraphicFrameLayout(dimension(199n,2n),100,new SourceAffineBudget()).widthEmu).toBe(100)
 expect(()=>projectSourceGraphicFrameLayout(dimension(1n<<513n),100,new SourceAffineBudget())).toThrow('precision')
})
