import {chartRational,chartRationalAdd,chartRationalCompare,chartRationalDivide,chartRationalMultiply,chartRationalSubtract,type ChartRational} from './chartRational.js'
export interface ChartRationalPoint {readonly x:ChartRational;readonly y:ChartRational}
export interface ClippedChartSegment {readonly start:ChartRationalPoint;readonly end:ChartRationalPoint;readonly startParameter:ChartRational;readonly endParameter:ChartRational}
const zero=chartRational(0n),one=chartRational(1n)

/** Intersect an oriented segment with [0,1]² using exact parametric slabs.
 * Preserve segment order, corners, and zero-length in-range segments. */
export function clipChartUnitSegment(start:ChartRationalPoint,end:ChartRationalPoint):ClippedChartSegment|undefined{
 let low=zero,high=one
 const dx=chartRationalSubtract(end.x,start.x),dy=chartRationalSubtract(end.y,start.y)
 for(const [origin,delta] of [[start.x,dx],[start.y,dy]]){
  if(delta!.numerator===0n){if(chartRationalCompare(origin!,zero)<0||chartRationalCompare(origin!,one)>0)return;continue}
  let entry=chartRationalDivide(chartRationalSubtract(zero,origin!),delta!),exit=chartRationalDivide(chartRationalSubtract(one,origin!),delta!)
  if(chartRationalCompare(entry,exit)>0)[entry,exit]=[exit,entry]
  if(chartRationalCompare(entry,low)>0)low=entry
  if(chartRationalCompare(exit,high)<0)high=exit
  if(chartRationalCompare(low,high)>0)return
 }
 const point=(parameter:ChartRational):ChartRationalPoint=>({x:chartRationalAdd(start.x,chartRationalMultiply(parameter,dx)),y:chartRationalAdd(start.y,chartRationalMultiply(parameter,dy))})
 return {start:point(low),end:point(high),startParameter:low,endParameter:high}
}
