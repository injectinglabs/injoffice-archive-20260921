import {describe,it,expect} from 'vitest'
import {createSignedChartStackBands} from './chartSignedStacking.js'

const input=(values:readonly string[])=>values.map((value,order)=>({index:100+order*3,order,values:[value]}))
const fraction=(n:bigint,d=1n)=>({numerator:n,denominator:d})
describe('Office-qualified signed Cartesian stacking',()=>{
 it('distinguishes diverging bars from algebraic lines in exact source order',()=>{
  const source=input(['4','-2','3','-1']),before=JSON.stringify(source)
  const bar=createSignedChartStackBands(source,'bar','stacked'),line=createSignedChartStackBands(source,'line','stacked')
  expect(bar.bands.map(b=>[b.lower[0],b.upper[0]])).toEqual([[0n,4n],[0n,-2n],[4n,7n],[-2n,-3n]].map(pair=>pair.map(n=>fraction(n))))
  expect(line.bands.map(b=>b.upper[0])).toEqual([4n,2n,5n,4n].map(n=>fraction(n)))
  expect(line.bands.map(b=>b.index)).toEqual([100,103,106,109])
 expect(JSON.stringify(source)).toBe(before)
 })
 it('uses XML array order while retaining distinct original c:order metadata',()=>{
  const source=input(['4','-2','3','-1']).reverse(),before=JSON.stringify(source)
  const line=createSignedChartStackBands(source,'line','percentStacked')
  expect(line.bands.map(b=>b.upper[0])).toEqual([fraction(-1n,10n),fraction(1n,5n),fraction(0n),fraction(2n,5n)])
  expect(line.bands.map(b=>b.order)).toEqual([3,2,1,0]);expect(JSON.stringify(source)).toBe(before)
  expect(createSignedChartStackBands(source,'bar','stacked').bands.map(b=>[b.lower[0],b.upper[0]])).toEqual([[0n,-1n],[0n,3n],[-1n,-3n],[3n,7n]].map(pair=>pair.map(n=>fraction(n))))
 })
 it('retains sign over absolute category totals, including cancellation and all-negative data',()=>{
  const series=[{index:7,order:0,values:['4','-2','2','0']},{index:2,order:1,values:['-2','-3','-2','-0']},{index:9,order:2,values:['3','0','0','0']},{index:5,order:3,values:['-1','0','0','0']}]
  const bar=createSignedChartStackBands(series,'bar','percentStacked'),line=createSignedChartStackBands(series,'line','percentStacked')
  expect(bar.totals).toEqual([10n,5n,4n,0n].map(n=>fraction(n)))
  expect(bar.bands[1]!.upper).toEqual([fraction(-1n,5n),fraction(-1n),fraction(-1n,2n),fraction(0n)])
  expect(line.bands[1]!.upper).toEqual([fraction(1n,5n),fraction(-1n),fraction(0n),fraction(0n)])
  expect(line.bands[3]!.upper[0]).toEqual(fraction(2n,5n))
 })
 it('preserves tiny and beyond-double exact contributions',()=>{
  for(const pair of [['1e-100','-2e-100'],['9007199254740993','-18014398509481986']]){
   const result=createSignedChartStackBands(input(pair),'line','percentStacked')
   expect(result.bands.map(b=>b.upper[0])).toEqual([fraction(1n,3n),fraction(-1n,3n)])
  }
 })
 it('rejects ambiguous order, sparse alignment, duplicate indices and resource overflow',()=>{
  const valid=input(['1','2'])
  for(const invalid of [[{...valid[0]!,order:-0}], [{...valid[0]!,index:-0}], [valid[0]!,{...valid[1]!,index:100}], [valid[0]!,{...valid[1]!,order:0}], [valid[0]!,{...valid[1]!,values:[]}],input(Array(17).fill('1')), [{index:0,order:0,values:Array(257).fill('1')}]])expect(()=>createSignedChartStackBands(invalid,'bar','stacked')).toThrow()
  const sparse=Array(2);sparse[0]=valid[0];expect(()=>createSignedChartStackBands(sparse,'bar','stacked')).toThrow()
 })
})
