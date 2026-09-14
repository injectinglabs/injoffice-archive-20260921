import {describe,expect,it} from 'vitest'
import {composeSourceAffines,convertSourceAffine,qualifySourceAffinePoint,sourceHierarchyAffine,SourceAffineBudget,type SourceAffineFrame} from './sourceAffine.js'
import {sourceBodyRotation,sourceUprightTextArea} from './sourceTextOrientation.js'

const frame:SourceAffineFrame={x:0,y:0,cx:400,cy:200}
function compile(leaf:SourceAffineFrame,rotation:number,parents:Parameters<typeof sourceHierarchyAffine>[1]=[]){
  const budget=new SourceAffineBudget()
  return composeSourceAffines(sourceHierarchyAffine(leaf,parents,budget),sourceBodyRotation(leaf,parents,rotation,budget),budget)
}
const numbers=(affine:ReturnType<typeof compile>)=>convertSourceAffine(affine).matrix
describe('source body rotation',()=>{
  it('matches the Microsoft +90 shape / −90 text example exactly on a non-square frame',()=>{
    const result=compile({...frame,rotation:5400000},-5400000)
    expect(numbers(result)).toEqual([1,0,0,1,0,0])
    expect(result.errors.every(error=>error.numerator===0n)).toBe(true)
  })
  it('rotates the text anchor after unequal group scales, with orthogonal glyph axes',()=>{
    const parents=[{frame:{x:0,y:0,cx:800,cy:200},child:frame}]
    // Group-scaled anchor center is (400,100). R90·diag(2,1)
    // has columns (0,2),(-1,0), and maps local center (200,100) there.
    expect(numbers(compile(frame,5400000,parents))).toEqual([0,2,-1,0,500,-300])
  })
  it('retains counter-reflection order for H, V and nested flips',()=>{
    expect(numbers(compile({...frame,flipH:true},5400000))).toEqual([0,1,-1,0,300,-100])
    expect(numbers(compile({...frame,flipV:true},5400000))).toEqual([0,-1,1,0,100,300])
    const parents=[{frame:{...frame,flipH:true},child:frame}]
    expect(numbers(compile({...frame,flipH:true},5400000,parents))).toEqual([0,1,-1,0,300,-100])
  })
  it('retains signed raw angle semantics while reducing complete turns only for evaluation',()=>{
    expect(numbers(compile(frame,21600000))).toEqual(numbers(compile(frame,0)))
    expect(numbers(compile(frame,-16200000))).toEqual(numbers(compile(frame,5400000)))
  })
  it('preserves noncentral text regions and requires the transformed ink/clip envelope to fit',()=>{
    const source=Object.freeze({...frame,x:1000,y:2000})
    const textRect=Object.freeze({x:35,y:60,cx:140,cy:50})
    const result=compile(source,5400000)
    // Rotation about whole frame center (1200,2100), not inset rectangle center.
    const [a,b,c,d,e,f]=numbers(result)
    expect([a*textRect.x+c*textRect.y+e,b*textRect.x+d*textRect.y+f]).toEqual([1240,1935])
    for(const [x,y] of [[35,60],[175,60],[175,110],[35,110]])qualifySourceAffinePoint(result,x!,y!,10000)
    expect(()=>qualifySourceAffinePoint(result,35,-10000,10000)).toThrow(/world coordinate/)
    expect(source).toEqual({...frame,x:1000,y:2000})
    expect(textRect).toEqual({x:35,y:60,cx:140,cy:50})
  })
  it('qualifies a noncardinal composition with the shared uncertainty budget',()=>{
    const result=compile({...frame,rotation:60000},120000)
    const [a,b,c,d]=numbers(result)
    expect(a).toBeCloseTo(Math.cos(Math.PI/60),12)
    expect(b).toBeCloseTo(Math.sin(Math.PI/60),12)
    expect(c).toBeCloseTo(-b,12);expect(d).toBeCloseTo(a,12)
    qualifySourceAffinePoint(result,400,200,10000)
  })
  it('refuses malformed angles, frames and exhausted request budgets',()=>{
    for(const angle of [-0,0.5,NaN,Infinity,2147483648,-2147483649])expect(()=>sourceBodyRotation(frame,[],angle,new SourceAffineBudget())).toThrow()
    expect(()=>sourceBodyRotation({...frame,cx:0},[],0,new SourceAffineBudget())).toThrow()
    const budget=new SourceAffineBudget();budget.charge(1_000_000)
    expect(()=>sourceBodyRotation(frame,[],0,budget)).toThrow(/budget/)
  })
})

describe('upright exact-area proposal',()=>{
  it('uses exact half-open quadrant boundaries',()=>{
    for(const [rotation,swapped] of [[0,false],[2699999,false],[2700000,true],[8099999,true],[8100000,false],[13499999,false],[13500000,true],[18899999,true],[18900000,false],[21599999,false]] as const){
      const result=sourceUprightTextArea({...frame,rotation},[],{x:20,y:30,cx:200,cy:100},new SourceAffineBudget())
      expect(result.swapped).toBe(swapped)
      expect(result.area.cx).toEqual({numerator:swapped?100n:200n,denominator:1n})
    }
  })
  it('keeps fractional area dimensions and asymmetric center without changing glyph scales',()=>{
    const leaf={...frame,rotation:5400000,flipV:true}
    const parents=[{frame:{...frame,cx:600},child:frame}]
    const content=Object.freeze({x:35,y:60,cx:150,cy:100})
    const budget=new SourceAffineBudget()
    const plan=sourceUprightTextArea(leaf,parents,content,budget)
    expect(plan.area).toEqual({x:{numerator:230n,denominator:3n},y:{numerator:-5n,denominator:2n},cx:{numerator:200n,denominator:3n},cy:{numerator:225n,denominator:1n}})
    // Glyph x scale remains 1.5, glyph y scale remains 1; swapping those
    // scales would make glyphs narrower/taller merely to avoid fractions.
    expect(numbers(composeSourceAffines(sourceHierarchyAffine(leaf,parents,budget),plan.orientation,budget))).toEqual([1.5,0,0,1,0,0])
    expect(content).toEqual({x:35,y:60,cx:150,cy:100})
  })
  it('uses accumulated rotation and flips without reinterpreting writing-mode or body rotation',()=>{
    const leaf={...frame,rotation:2700000,flipH:true}
    const parents=[{frame:{...frame,rotation:2700000,flipV:true},child:frame}]
    const budget=new SourceAffineBudget()
    const plan=sourceUprightTextArea(leaf,parents,{x:0,y:0,cx:400,cy:200},budget)
    expect(plan.swapped).toBe(true)
    const world=composeSourceAffines(sourceHierarchyAffine(leaf,parents,budget),plan.orientation,budget)
    const [a,b,c,d]=numbers(world)
    expect([a,b,c,d]).toEqual([1,0,0,1])
    qualifySourceAffinePoint(world,400,200,10000,budget)
  })
  it('refuses collapsed or noninteger source content and exhausted budgets',()=>{
    for(const content of [{x:0,y:0,cx:0,cy:1},{x:0.5,y:0,cx:1,cy:1}])expect(()=>sourceUprightTextArea(frame,[],content,new SourceAffineBudget())).toThrow()
    const budget=new SourceAffineBudget();budget.charge(1_000_000)
    expect(()=>sourceUprightTextArea(frame,[],{x:0,y:0,cx:1,cy:1},budget)).toThrow(/budget/)
  })
})
