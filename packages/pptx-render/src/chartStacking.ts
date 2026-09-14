import {chartRational,chartRationalAdd,chartRationalDecimal,chartRationalDivide,type ChartRational} from './chartRational.js'

export type ChartStackGrouping='standard'|'stacked'|'percentStacked'
export interface ChartStackSeries {readonly index:number;readonly order:number;readonly values:readonly string[]}
export interface ChartStackBand {readonly index:number;readonly order:number;readonly lower:readonly ChartRational[];readonly upper:readonly ChartRational[]}
export interface ChartStackResult {readonly bands:readonly ChartStackBand[];readonly totals:readonly ChartRational[]}

/** Complete aligned category data, in authored series order. Signed standard
 * values are supported; negative stacked values remain unqualified until the
 * family-specific Office accumulation/percentage rules have reader evidence.
 * Percent values use a unit scale (1 = 100%), never rounded source decimals. */
export function createChartStackBands(series:readonly ChartStackSeries[],grouping:ChartStackGrouping):ChartStackResult{
 if(!['standard','stacked','percentStacked'].includes(grouping)||series.length<1||series.length>16)throw new RangeError('invalid chart stack profile')
 for(let i=0;i<series.length;i++)if(!Object.hasOwn(series,i)||!series[i])throw new RangeError('sparse chart stack series')
 const count=series[0]!.values.length
 if(count<1||count>256)throw new RangeError('invalid chart stack point count')
 const indices=new Set<number>(),orders=new Set<number>()
 const values=series.map((item,order)=>{
  if(!Number.isInteger(item.index)||Object.is(item.index,-0)||item.index<0||item.index>4294967295||indices.has(item.index)||!Number.isSafeInteger(item.order)||Object.is(item.order,-0)||item.order<0||item.order>=series.length||orders.has(item.order)||item.values.length!==count)throw new RangeError('invalid chart stack series alignment')
  indices.add(item.index);orders.add(item.order)
  return item.values.map(raw=>{
   const value=chartRationalDecimal(raw)
   if(grouping!=='standard'&&value.numerator<0n)throw new RangeError('negative chart stacking is not qualified')
   return value
  })
 })
 const zero=chartRational(0n),totals=Array.from({length:count},(_,point)=>values.reduce((sum,row)=>chartRationalAdd(sum,row[point]!),zero))
 let accumulated=Array.from({length:count},()=>zero)
 const bands=series.map((item,order)=>{
  const lower=grouping==='standard'?Array.from({length:count},()=>zero):accumulated
  const upper=values[order]!.map((value,point)=>chartRationalAdd(lower[point]!,value))
  accumulated=upper
  const scale=(boundary:readonly ChartRational[])=>grouping==='percentStacked'?boundary.map((value,point)=>totals[point]!.numerator===0n?zero:chartRationalDivide(value,totals[point]!)):boundary
  return {index:item.index,order:item.order,lower:scale(lower),upper:scale(upper)}
 })
 return {bands,totals}
}
