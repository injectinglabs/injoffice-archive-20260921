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
 // Outer scale gives group center (12,2). Its quarter quadrant swaps
 // inherited scales, so local (1,1) maps to (15,0) with physical size (4,6).
 expect([value(result.width),value(result.height)]).toEqual([4,6])
 expect(result.origin.values.map(value)).toEqual([1,0,0,1,13,-3])
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

// Original source-model coordinates differ from Office's import rewrite for
// two noncardinal nested tables. Preserve both observations, without treating
// Office-import parity as an acceptance claim for this source-anchored policy.
it.each([
 ['quarter',5400000,0,false,3190476.1904761903,2288095.2380952383],
 ['turn-reflect',1800000,0,true,2530162.6940632984,2307646.4862549203],
 ['double-turn',1800000,1800000,false,3295924.942316509,2512831.862268609],
] as const)('preserves original source-anchored nested %s layout',(_name,innerRotation,outerRotation,outerFlip,x,y)=>{
 const result=sourceGraphicFrameLayout({x:500000,y:200000,cx:4000000,cy:2000000},[
  {frame:{x:3000000,y:2200000,cx:6000000,cy:2000000,rotation:innerRotation},child:{x:0,y:0,cx:4000000,cy:2000000}},
  {frame:{x:500000,y:300000,cx:10000000,cy:5000000,rotation:outerRotation,flipH:outerFlip},child:{x:0,y:0,cx:12000000,cy:7000000}},
 ],new SourceAffineBudget())
 expect(value(result.origin.values[4])).toBeCloseTo(x,7)
 expect(value(result.origin.values[5])).toBeCloseTo(y,7)
})
it.each([
 ['quarter',5400000,0,false,251.21875,180.164062,337.457031,131.234376],
 ['turn-reflect',1800000,0,true,199.226562,181.703125,393.699219,112.488281],
 ['double-turn',1800000,1800000,false,259.523438,197.859375,393.699218,112.488281],
] as const)('matches retained filled-chart nested %s physical anchor and size',(_name,innerRotation,outerRotation,outerFlip,x,y,w,h)=>{
 const result=sourceGraphicFrameLayout({x:500000,y:200000,cx:4000000,cy:2000000},[
  {frame:{x:3000000,y:2200000,cx:6000000,cy:2000000,rotation:innerRotation},child:{x:0,y:0,cx:4000000,cy:2000000}},
  {frame:{x:500000,y:300000,cx:10000000,cy:5000000,rotation:outerRotation,flipH:outerFlip},child:{x:0,y:0,cx:12000000,cy:7000000}},
 ],new SourceAffineBudget())
 for(const [actual,reference] of [[value(result.origin.values[4]),x],[value(result.origin.values[5]),y],[value(result.origin.values[4])+value(result.width),x+w],[value(result.origin.values[5])+value(result.height),y+h]])expect(Math.abs(actual!-reference!*12700)).toBeLessThan(25)
})
it('counts source hierarchy depth independently from bounded arithmetic compositions',()=>{
 const ancestor={frame:{x:0,y:0,cx:2,cy:2},child:{x:0,y:0,cx:2,cy:2}}
 const result=sourceGraphicFrameLayout({x:1,y:1,cx:1,cy:1},Array.from({length:63},()=>ancestor),new SourceAffineBudget())
 expect(result.origin.depth).toBe(64);expect(result.origin.values.map(value)).toEqual([1,0,0,1,1,1])
 expect(()=>sourceGraphicFrameLayout({x:0,y:0,cx:1,cy:1},Array.from({length:64},()=>ancestor),new SourceAffineBudget())).toThrow('depth')
})

// Exact transforms retained from Office-normalized source; original source
// import rewrite is separate retained comparison evidence, not emulated here.
const normalizedOfficeTables=[
 {
  "name": "table-nested-quarter",
  "leaf": {
   "x": 788888,
   "y": 171430,
   "cx": 3733333,
   "cy": 2400000
  },
  "parents": [
   {
    "frame": {
     "x": 2828571,
     "y": 2983333,
     "cx": 5600000,
     "cy": 2400000,
     "rotation": 5400000
    },
    "child": {
     "x": 788888,
     "y": 171430,
     "cx": 3733333,
     "cy": 2400000
    }
   },
   {
    "frame": {
     "x": 4190476,
     "y": 1288095,
     "cx": 2000000,
     "cy": 4000000
    },
    "child": {
     "x": 4428571,
     "y": 1383333,
     "cx": 2400000,
     "cy": 5600000
    }
   }
  ],
  "reference": [
   3190478.125,
   2288083.5874
  ]
 },
 {
  "name": "table-nested-turn-reflect",
  "leaf": {
   "x": 1360696,
   "y": -203590,
   "cx": 3200000,
   "cy": 2800000
  },
  "parents": [
   {
    "frame": {
     "x": 4749775,
     "y": 2690618,
     "cx": 4800000,
     "cy": 2800000,
     "rotation": 1800000
    },
    "child": {
     "x": 1360696,
     "y": -203590,
     "cx": 3200000,
     "cy": 2800000
    }
   },
   {
    "frame": {
     "x": 2541854,
     "y": 2221870,
     "cx": 4000000,
     "cy": 2000000,
     "flipH": true
    },
    "child": {
     "x": 4749775,
     "y": 2690618,
     "cx": 4800000,
     "cy": 2800000
    }
   }
  ],
  "reference": [
   2541835.5437,
   2221855.0813
  ]
 },
 {
  "name": "table-nested-double-turn",
  "leaf": {
   "x": 897948,
   "y": 606218,
   "cx": 3200000,
   "cy": 2800000
  },
  "parents": [
   {
    "frame": {
     "x": 3743744,
     "y": 3044871,
     "cx": 4800000,
     "cy": 2800000,
     "rotation": 1800000
    },
    "child": {
     "x": 897948,
     "y": 606218,
     "cx": 3200000,
     "cy": 2800000
    }
   },
   {
    "frame": {
     "x": 3266284,
     "y": 2444381,
     "cx": 4000000,
     "cy": 2000000,
     "rotation": 1800000
    },
    "child": {
     "x": 3743744,
     "y": 3044871,
     "cx": 4800000,
     "cy": 2800000
    }
   }
  ],
  "reference": [
   3266281.25,
   2444402.7312
  ]
 }
]
it.each(normalizedOfficeTables)('matches retained normalized source $name',({leaf,parents,reference})=>{
 const result=sourceGraphicFrameLayout(leaf,parents,new SourceAffineBudget())
 expect(Math.abs(value(result.origin.values[4])-reference[0]!)).toBeLessThan(25)
 expect(Math.abs(value(result.origin.values[5])-reference[1]!)).toBeLessThan(25)
})
