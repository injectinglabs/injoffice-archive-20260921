import {describe,it,expect} from 'vitest'
import {createChartStackBands} from './chartStacking.js'
const series=(values:string[][])=>values.map((row,order)=>({index:10+order,order,values:row}))
const text=(value:{numerator:bigint;denominator:bigint})=>`${value.numerator}/${value.denominator}`
describe('exact chart stack boundaries',()=>{
 it('aligns points and authored series order without mutating source lexemes',()=>{
  const source=series([['+.10','2e1'],['.20','3'],['.30','4']]),before=JSON.stringify(source)
  const result=createChartStackBands(source,'stacked')
  expect(result.bands.map(b=>b.lower.map(text))).toEqual([['0/1','0/1'],['1/10','20/1'],['3/10','23/1']])
  expect(result.bands[2]!.upper.map(text)).toEqual(['3/5','27/1'])
  expect(JSON.stringify(source)).toBe(before)
 })
 it('uses exact category totals and unit percentages including repeating thirds',()=>{
  const result=createChartStackBands(series([['1','1e100'],['2','2e100']]),'percentStacked')
  expect(result.bands[0]!.upper.map(text)).toEqual(['1/3','1/3'])
  expect(result.bands[1]!.lower.map(text)).toEqual(['1/3','1/3'])
  expect(result.bands[1]!.upper.map(text)).toEqual(['1/1','1/1'])
 })
 it('keeps all zero categories degenerate, including negative zero',()=>{
  const result=createChartStackBands(series([['-0','0'],['0e99','+0']]),'percentStacked')
  expect(result.bands.every(b=>[...b.lower,...b.upper].every(v=>v.numerator===0n))).toBe(true)
 })
 it('uses an independent zero baseline for signed standard area',()=>{
  const result=createChartStackBands(series([['-3','2'],['1','-4']]),'standard')
  expect(result.bands[1]!.lower.map(text)).toEqual(['0/1','0/1'])
  expect(result.bands[1]!.upper.map(text)).toEqual(['1/1','-4/1'])
 })
 it('refuses unqualified negative stacks and malformed alignment before output',()=>{
  for(const grouping of ['stacked','percentStacked'] as const)expect(()=>createChartStackBands(series([['-1']]),grouping)).toThrow(/negative/)
  for(const source of [[],series([[]]),series([['1'],['2','3']]),[{index:0,order:1,values:['1']}],[{index:0,order:0,values:['1']},{index:0,order:1,values:['2']}],series([['NaN']])])expect(()=>createChartStackBands(source,'stacked')).toThrow()
 })
})
