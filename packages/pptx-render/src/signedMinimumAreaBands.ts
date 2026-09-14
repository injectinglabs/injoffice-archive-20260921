import {chartRationalDecimal as decimal,chartRationalCompare as compare} from './chartRational.js'
import {createSignedChartStackBands} from './chartSignedStacking.js'
import type {ChartStackGrouping,ChartStackSeries,ChartStackResult} from './chartStacking.js'
/** Source minimum changes only first closure. Area prefixes are algebraic,
 * percentage denominators are absolute sums; no split-sign bar accumulation. */
export function createSignedMinimumAreaBands(series:readonly ChartStackSeries[],grouping:ChartStackGrouping,crossing:'min',minimum:string,maximum:string):ChartStackResult{
 if(crossing!=='min'||!['stacked','percentStacked'].includes(grouping)||!Array.isArray(series)||series.length<1||series.length>16)throw new RangeError('invalid signed minimum area profile')
 for(let i=0;i<series.length;i++)if(!Object.hasOwn(series,i)||!series[i]||!Array.isArray(series[i]!.values))throw new RangeError('invalid signed area series')
 const baseline=decimal(minimum),high=decimal(maximum)
 if(compare(baseline,high)>=0||baseline.numerator>0n||high.numerator<0n)throw new RangeError('invalid explicit signed area scale')
 const stack=createSignedChartStackBands(series,'line',grouping==='percentStacked'?'percentStacked':'stacked')
 return {totals:stack.totals,bands:stack.bands.map((band,index)=>index===0?{...band,lower:band.lower.map(()=>baseline)}:band)}
}
