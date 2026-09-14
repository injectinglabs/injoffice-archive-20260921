import {expect,it} from 'vitest'
import {sourceAffine,sourceHierarchyAffine,composeSourceAffines,convertSourceAffine,qualifySourceAffinePoint,identitySourceAffine,encodeSourceAffine,decodeSourceAffine,SourceAffineBudget,SOURCE_AFFINE_LIMITS} from './sourceAffine.js'
const numeric=(a:ReturnType<typeof sourceAffine>)=>convertSourceAffine(a).matrix
it('retains exact cardinal/reflection center translations including half EMU',()=>{
 expect(numeric(sourceAffine({x:100,y:200,cx:3,cy:2,rotation:5400000}))).toEqual([0,1,-1,0,102.5,199.5])
 const flip=sourceAffine({x:100,y:200,cx:3,cy:2,flipH:true})
 expect(numeric(flip)).toEqual([-1,0,0,1,103,200]);expect(flip.errors.every(r=>r.numerator===0n)).toBe(true)
})
it('applies AnnexL flip before rotation, not their noncommuting reverse',()=>{
 expect(numeric(sourceAffine({x:10,y:20,cx:8,cy:4,rotation:5400000,flipH:true}))).toEqual([0,-1,-1,0,16,26])
})
it('maps child center through exact nonuniform group scale before rotation',()=>{
 const affine=sourceAffine({x:10,y:20,cx:8,cy:6,rotation:5400000},{x:2,y:3,cx:4,cy:2})
 expect(numeric(affine)).toEqual([0,2,-3,0,26,15])
 const composed=composeSourceAffines(affine,sourceAffine({x:3,y:4,cx:1,cy:1}))
 expect(numeric(composed)).toEqual([0,2,-3,0,14,21])
})
it('uses DrawingML leaf-axis scaling rather than conventional nesting for rotated children',()=>{
 const leaf={x:0,y:0,cx:2,cy:2,rotation:5400000},parent={frame:{x:0,y:0,cx:4,cy:2},child:{x:0,y:0,cx:2,cy:2}}
 const conventional=composeSourceAffines(sourceAffine(parent.frame,parent.child),sourceAffine(leaf))
 expect(numeric(conventional)).toEqual([0,1,-2,0,4,0])
 // AnnexL: center maps to(2,1), own horizontal scale2 then90degree turn.
 expect(numeric(sourceHierarchyAffine(leaf,[parent]))).toEqual([0,2,-1,0,3,-1])
})
it('sums nested rotations and multiplies own-axis flips while mapping centers conventionally',()=>{
 const leaf={x:2,y:3,cx:4,cy:2,rotation:5400000,flipH:true}
 const parent={frame:{x:10,y:20,cx:8,cy:6,rotation:5400000,flipH:true},child:{x:2,y:3,cx:4,cy:2}}
 // Both horizontal flips cancel; total rotation180, scales2 and3.
 expect(numeric(sourceHierarchyAffine(leaf,[parent]))).toEqual([-2,0,0,-3,18,26])
})
it('matches independent 30degree corner equations within retained uncertainty',()=>{
 const affine=sourceAffine({x:0,y:0,cx:4000000,cy:2000000,rotation:1800000})
 const converted=convertSourceAffine(affine),m=converted.matrix
 expect(m[0]).toBeCloseTo(Math.sqrt(3)/2,14);expect(m[1]).toBeCloseTo(0.5,14)
 expect(m[4]).toBeCloseTo(2500000-1000000*Math.sqrt(3),7)
 expect(m[5]).toBeCloseTo(-500000*Math.sqrt(3),7)
 for(const [x,y] of [[0,0],[4000000,0],[0,2000000],[4000000,2000000]])expect(()=>qualifySourceAffinePoint(converted.qualified,x!,y!,281474976710655)).not.toThrow()
})
it('keeps large translation cancellation exact before paint conversion',()=>{
 const forward=sourceAffine({x:9007199254740990,y:0,cx:1,cy:1}),reverse=sourceAffine({x:-9007199254740989,y:0,cx:1,cy:1})
 expect(numeric(composeSourceAffines(forward,reverse))).toEqual([1,0,0,1,1,0])
})
it('refuses cumulative error growth, excessive depth and world bounds',()=>{
 const huge=convertSourceAffine(sourceAffine({x:0,y:0,cx:100000000000000,cy:100000000000000,rotation:2700000})).qualified
 expect(()=>qualifySourceAffinePoint(huge,100000000000000,0,281474976710655)).toThrow(/uncertainty/)
 expect(()=>qualifySourceAffinePoint(identitySourceAffine(),1001,0,1000)).toThrow(/coordinate/)
 let affine=identitySourceAffine();const move=sourceAffine({x:1,y:0,cx:1,cy:1})
 for(let i=0;i<64;i++)affine=composeSourceAffines(affine,move)
 expect(()=>composeSourceAffines(affine,move)).toThrow(/depth/)
})
it('roundtrips canonical bounded rational transport and rejects ambiguous tuples',()=>{
 const source=sourceAffine({x:0,y:0,cx:123456,cy:654321,rotation:1}),wire=encodeSourceAffine(source)
 expect(decodeSourceAffine(JSON.parse(JSON.stringify(wire)))).toEqual(source)
 for(const bad of [
  {...wire,values:Array(6)},
  {...wire,values:[['2','2'],...wire.values.slice(1)]},
  {...wire,values:[['-0','1'],...wire.values.slice(1)]},
  {...wire,errors:[['-1','1'],...wire.errors.slice(1)]},
  {...wire,values:[[(1n<<512n).toString(),'1'],...wire.values.slice(1)]},
  {...wire,depth:65},
  {...wire,extra:0},
 ])expect(()=>decodeSourceAffine(bad)).toThrow()
 let calls=0;expect(()=>decodeSourceAffine({...wire,get values(){calls++;return wire.values}})).toThrow();expect(calls).toBe(0)
})
it('enforces shared operation and conservative transport-byte budgets',()=>{
 const operations=new SourceAffineBudget();operations.charge(SOURCE_AFFINE_LIMITS.maxOperations)
 expect(()=>sourceAffine({x:0,y:0,cx:1,cy:1},undefined,operations)).toThrow(/Aggregate/)
 const payload=new SourceAffineBudget();payload.charge(0,SOURCE_AFFINE_LIMITS.maxPayloadBytes)
 expect(()=>encodeSourceAffine(identitySourceAffine(),payload)).toThrow(/Aggregate/)
})
