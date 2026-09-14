import {chartRational,chartRationalAdd,chartRationalDecimal,chartRationalDivide,type ChartRational} from './chartRational.js'
import type {ChartStackSeries,ChartStackResult} from './chartStacking.js'

export type SignedStackFamily='bar'|'line'
export type SignedStackGrouping='stacked'|'percentStacked'

/** Dense aligned series in XML array order, retaining the original c:order
 * permutation as metadata. Verified against the isolated Office
 * 16.112.4 sign matrix: bars accumulate each sign separately; lines accumulate
 * algebraically. Both percentage families divide by the category absolute sum,
 * retaining signs. A zero absolute sum produces zero boundaries, never NaN.
 * Axis orientation, clipping and paint order belong to the family geometry. */
export function createSignedChartStackBands(series:readonly ChartStackSeries[],family:SignedStackFamily,grouping:SignedStackGrouping):ChartStackResult {
 if(!['bar','line'].includes(family)||!['stacked','percentStacked'].includes(grouping)||series.length<1||series.length>16)throw new RangeError('invalid signed chart stack profile')
 const count=series[0]!.values.length
 if(count<1||count>256)throw new RangeError('invalid signed chart stack point count')
 for(let i=0;i<series.length;i++)if(!Object.hasOwn(series,i))throw new RangeError('sparse signed chart stack series')
 const indices=new Set<number>(),orders=new Set<number>()
 const values=series.map(item=>{
  if(!Number.isInteger(item.index)||Object.is(item.index,-0)||item.index<0||item.index>4294967295||indices.has(item.index)||!Number.isInteger(item.order)||Object.is(item.order,-0)||item.order<0||item.order>=series.length||orders.has(item.order)||item.values.length!==count)throw new RangeError('invalid signed chart stack series alignment')
  indices.add(item.index);orders.add(item.order)
  for(let i=0;i<count;i++)if(!Object.hasOwn(item.values,i))throw new RangeError('sparse signed chart stack values')
  return item.values.map(chartRationalDecimal)
 })
 const zero=chartRational(0n)
 const totals=Array.from({length:count},(_,point)=>values.reduce((sum,row)=>{
  const value=row[point]!
  return chartRationalAdd(sum,value.numerator<0n?chartRational(-value.numerator,value.denominator):value)
 },zero))
 const positive=Array.from({length:count},()=>zero),negative=Array.from({length:count},()=>zero),running=Array.from({length:count},()=>zero)
 const bands=series.map((item,order)=>{
  const lower:ChartRational[]=[],upper:ChartRational[]=[]
  for(let point=0;point<count;point++){
   const value=values[order]![point]!
   const accumulator=family==='line'?running:value.numerator<0n?negative:positive
   const from=accumulator[point]!,to=chartRationalAdd(from,value)
   accumulator[point]=to
   const scale=(boundary:ChartRational)=>grouping==='percentStacked'?(totals[point]!.numerator===0n?zero:chartRationalDivide(boundary,totals[point]!)):boundary
   lower.push(scale(from));upper.push(scale(to))
  }
  return {index:item.index,order:item.order,lower,upper}
 })
 return {bands,totals}
}
