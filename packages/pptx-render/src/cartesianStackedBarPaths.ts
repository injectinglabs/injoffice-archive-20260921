import type {NativeLiteralBar} from '@injoffice/pptx-native'
import {createCartesianBarPaths,type CartesianBarVector} from './cartesianBarPaths.js'
import {createSignedChartStackBands,type SignedStackGrouping} from './chartSignedStacking.js'
import {validateStackedChartGeometry} from './chartStackedGeometryValidation.js'
import {chartRational as r,chartRationalDecimal as decimal,chartRationalSubtract as sub,chartRationalDivide as div,chartRationalCompare as cmp,chartRationalCoordinate as coordinate,type ChartRational} from './chartRational.js'

export type CartesianStackedBarInput=Omit<NativeLiteralBar,'profile'|'dataOrigin'|'grouping'|'overlap'>&{readonly grouping:SignedStackGrouping;readonly overlap:100}
/** Fully overlapping stacked bars; source paint, point and series order survive.
 * Clip exact signed cumulative boundaries before the sole EMU quantization. */
export function createCartesianStackedBarPaths(chart:CartesianStackedBarInput,cx:number,cy:number):readonly CartesianBarVector[]{
 validateStackedChartGeometry(chart,'bar')
 if(chart.overlap!==100||!chart.series.length)throw new RangeError('stacked bars require full overlap')
 const bands=createSignedChartStackBands(chart.series,'bar',chart.grouping).bands
 if(bands[0]!.lower.length!==chart.categories.length)throw new RangeError('stacked bar categories and values differ')
 // Reuse the unchanged clustered axis/category/frame admission with one lane:
 // stack width is independent of series count. Its zero bars are discarded.
 const scaffold=createCartesianBarPaths({...chart,grouping:'clustered',overlap:0,series:[{...chart.series[0]!,order:0,values:chart.categories.map(()=> '0')}]},cx,cy)
 const column=chart.barDirection==='column',categoryExtent=column?cx:cy,valueExtent=column?cy:cx
 const categoryReverse=column?chart.categoryAxis.orientation==='maxMin':chart.categoryAxis.orientation==='minMax'
 const valueReverse=column?chart.valueAxis.orientation==='minMax':chart.valueAxis.orientation==='maxMin'
 const min=decimal(chart.valueAxis.min!),max=decimal(chart.valueAxis.max!),span=sub(max,min)
 const scale=(value:ChartRational)=>coordinate(div(sub(cmp(value,min)<0?min:cmp(value,max)>0?max:value,min),span),valueExtent,valueReverse)
 const denominator=BigInt(2*chart.categories.length*(100+chart.gapWidth))
 const category=(numerator:bigint)=>coordinate(r(numerator,denominator),categoryExtent,categoryReverse)
 const vectors:CartesianBarVector[]=[]
 for(const [order,series] of chart.series.entries()){
  if(series.colors.length!==chart.categories.length||(series.title!==undefined&&(typeof series.title!=='string'||series.title.length>1024)))throw new RangeError('invalid stacked bar series paint')
  for(let point=0;point<chart.categories.length;point++){
   const color=series.colors[point]!
   if(!/^#[0-9A-F]{6}$/.test(color))throw new RangeError('invalid stacked bar color')
   const numerator=BigInt(2*point*(100+chart.gapWidth)+chart.gapWidth)
   const a=category(numerator),b=category(numerator+200n),from=scale(bands[order]!.lower[point]!),to=scale(bands[order]!.upper[point]!)
   if(from===to)continue
   const low=Math.min(a,b),high=Math.max(a,b),near=Math.min(from,to),far=Math.max(from,to)
   const x=column?low:near,y=column?near:low,right=column?high:far,bottom=column?far:high
   vectors.push({seriesIndex:series.index,pointIndex:point,color,path:[{kind:'moveTo',x,y},{kind:'lineTo',x:right,y},{kind:'lineTo',x:right,y:bottom},{kind:'lineTo',x,y:bottom},{kind:'close'}]})
  }
 }
 return [...vectors,...scaffold.filter(vector=>vector.axis!==undefined)]
}
