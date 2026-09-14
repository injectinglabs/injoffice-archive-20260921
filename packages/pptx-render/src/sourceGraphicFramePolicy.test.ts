import {expect,it} from 'vitest'
import {SourceAffineBudget,sourceHierarchyAffine} from './sourceAffine.js'
import {sourceRenderTransform} from './sourceRenderTransform.js'
import {qualifySourceGraphicFrameSpans,sourceGraphicFramePaintTransform} from './sourceGraphicFramePolicy.js'
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
