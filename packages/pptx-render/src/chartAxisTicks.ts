import {chartRational as rational,chartRationalDecimal as decimal,chartRationalAdd as add,chartRationalSubtract as subtract,chartRationalDivide as divide,chartRationalCompare as compare,type ChartRational} from './chartRational.js'
import {formatChartFixedDecimal} from './chartNumberFormat.js'
export interface ChartAxisTick {readonly value:ChartRational;readonly position:ChartRational;readonly label:string}
/** Explicit min + k*majorUnit grid, with exact min/grid alignment. */
export function createChartAxisTicks(minimum:string,maximum:string,majorUnit:string,format:string):readonly ChartAxisTick[]{
 const low=decimal(minimum),high=decimal(maximum),step=decimal(majorUnit),zero=rational(0n)
 if(compare(low,high)>=0||compare(step,zero)<=0||divide(low,step).denominator!==1n)throw new RangeError('invalid explicit chart tick grid')
 const span=subtract(high,low),ratio=divide(span,step),count=ratio.numerator/ratio.denominator+1n
 if(count>256n)throw new RangeError('chart tick count budget exceeded')
 const ticks:ChartAxisTick[]=[]
 for(let value=low;compare(value,high)<=0;value=add(value,step)){
  if(ticks.length>=256)throw new RangeError('chart tick count budget exceeded')
  ticks.push({value,position:divide(subtract(value,low),span),label:formatChartFixedDecimal(value,format)})
 }
 return ticks
}
