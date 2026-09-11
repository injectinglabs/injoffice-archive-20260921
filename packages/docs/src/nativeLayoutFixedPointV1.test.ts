import { describe,it,expect } from 'vitest'
import { solveNativeDocxLayoutFixedPointV1 } from './nativeLayoutFixedPointV1.js'

describe('bounded combined native layout solve',()=>{
 it('requires both independent layout effects to stabilize',async()=>{
  const value=await solveNativeDocxLayoutFixedPointV1({field:1,wrap:0},async state=>({next:{field:2,wrap:state.field===2?5:0},result:state}))
  expect(value).toEqual({state:{field:2,wrap:5},result:{field:2,wrap:5},passes:3})
 })
 it('refuses a cycle instead of returning stale field values',async()=>{
  await expect(solveNativeDocxLayoutFixedPointV1(0,async state=>({next:1-state,result:state}))).rejects.toThrow('cycle')
 })
 it('refuses nonconvergence after exactly eight bounded calls',async()=>{
  let calls=0
  await expect(solveNativeDocxLayoutFixedPointV1(0,async state=>{calls++;return {next:state+1,result:state}})).rejects.toThrow('eight passes')
  expect(calls).toBe(8)
 })
 it('canonicalizes key order and isolates callback mutation',async()=>{
  const seed={a:1,b:2}
  const result=await solveNativeDocxLayoutFixedPointV1(seed,async state=>{state.a=9;return {next:{b:2,a:1},result:'ok'}})
  expect(seed.a).toBe(1);expect(result.passes).toBe(1)
 })
})
