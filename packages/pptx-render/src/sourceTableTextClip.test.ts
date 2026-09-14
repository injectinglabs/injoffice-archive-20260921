import {describe,expect,it} from 'vitest'
import {identitySourceAffine,sourceAffine,SourceAffineBudget} from './sourceAffine.js'
import {sourceTableTextClip} from './sourceTableTextClip.js'
describe('finite table horizontal text clip',()=>{
  it('retains vertical overflow without extending the cell horizontal edges',()=>{
    const clip=sourceTableTextClip(100,[{bounds:{x:-20,y:-50,cx:180,cy:300},transform:identitySourceAffine()}],10000,new SourceAffineBudget())
    expect(clip).toEqual({x:0,y:-51,cx:100,cy:302})
    expect(clip.y).toBeLessThan(-50);expect(clip.y+clip.cy).toBeGreaterThan(250)
  })
  it('encloses rotated text and a separately untransformed refusal placeholder',()=>{
    const budget=new SourceAffineBudget()
    const rotated=sourceAffine({x:0,y:0,cx:100,cy:50,rotation:5400000,flipH:true},undefined,budget)
    const clip=sourceTableTextClip(100,[{bounds:{x:0,y:0,cx:100,cy:50},transform:rotated},{bounds:{x:0,y:-80,cx:30,cy:10},transform:identitySourceAffine()}],10000,budget)
    expect(clip).toEqual({x:0,y:-81,cx:100,cy:157})
  })
  it('qualifies uncertain local transforms and refuses clipped coordinate overflow',()=>{
    const budget=new SourceAffineBudget()
    const affine=sourceAffine({x:0,y:0,cx:100,cy:50,rotation:60000},undefined,budget)
    const clip=sourceTableTextClip(100,[{bounds:{x:0,y:0,cx:100,cy:50},transform:affine}],10000,budget)
    expect(clip.y).toBeLessThan(0);expect(clip.y+clip.cy).toBeGreaterThan(50)
    expect(()=>sourceTableTextClip(100,[{bounds:{x:0,y:0,cx:100,cy:10000},transform:identitySourceAffine()}],10000,new SourceAffineBudget())).toThrow(/coordinate/)
  })
  it('refuses missing/invalid hulls and exhausted aggregate budgets',()=>{
    expect(()=>sourceTableTextClip(100,[],10000,new SourceAffineBudget())).toThrow()
    expect(()=>sourceTableTextClip(100,[{bounds:{x:0,y:0,cx:-1,cy:1},transform:identitySourceAffine()}],10000,new SourceAffineBudget())).toThrow()
    const budget=new SourceAffineBudget();budget.charge(1_000_000)
    expect(()=>sourceTableTextClip(100,[{bounds:{x:0,y:0,cx:1,cy:1},transform:identitySourceAffine()}],10000,budget)).toThrow(/budget/)
  })
})
