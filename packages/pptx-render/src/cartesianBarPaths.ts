import {chartDecimalCoordinate,compareChartDecimals,parseChartDecimal} from './chartDecimal.js'
import type {RenderPathCommand,RenderStroke} from './types.js'

import type {NativeLiteralBar as LiteralBarRecord} from '@injoffice/pptx-native'
export interface CartesianBarVector {seriesIndex?:number;pointIndex?:number;axis?:'category'|'value';path:readonly RenderPathCommand[];color?:string;stroke?:RenderStroke}
const rgb=/^#[0-9A-F]{6}$/
function roundRatio(numerator:bigint,denominator:bigint):number{return Number((2n*numerator+denominator)/(2n*denominator))}

/** Explicit linear scale and clustered source data fitted to the full host frame.
 * Integer EMU quantization happens once per endpoint, after exact decimal/rational
 * arithmetic. Categories and series titles are metadata, not fabricated labels. */
export function createCartesianBarPaths(chart:Omit<LiteralBarRecord,'profile'|'dataOrigin'>,cx:number,cy:number):readonly CartesianBarVector[]{
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>281474976710655||cy>281474976710655)throw new RangeError('invalid literal bar frame')
 if(chart.grouping!=='clustered'||!['column','bar'].includes(chart.barDirection)||!Number.isInteger(chart.gapWidth)||chart.gapWidth<0||chart.gapWidth>500||chart.overlap!==0||chart.categories.length<1||chart.categories.length>256||chart.categories.some(c=>typeof c!=='string')||chart.categories.reduce((sum,c)=>sum+c.length,0)>32768||chart.series.length<1||chart.series.length>16)throw new RangeError('invalid literal bar profile')
 const category=chart.categoryAxis,value=chart.valueAxis,column=chart.barDirection==='column'
 for(const axis of [category,value]){
  if(!Number.isInteger(axis.id)||axis.id<0||axis.id>4294967295||!Number.isInteger(axis.crossAxisId)||axis.crossAxisId<0||axis.crossAxisId>4294967295||!['minMax','maxMin'].includes(axis.orientation)||typeof axis.deleted!=='boolean'||(axis.deleted?(axis.color!==undefined||axis.widthEmu!==undefined):(!rgb.test(axis.color??'')||!Number.isInteger(axis.widthEmu)||axis.widthEmu!<1||axis.widthEmu!>20116800)))throw new RangeError('invalid literal bar axis')
 }
 if(category.id===value.id||category.crossAxisId!==value.id||value.crossAxisId!==category.id||category.position!==(column?'b':'l')||value.position!==(column?'l':'b')||category.min!==undefined||category.max!==undefined||category.crossesAt!==undefined)throw new RangeError('invalid literal bar axis pairing')
 const minimum=parseChartDecimal(value.min!),maximum=parseChartDecimal(value.max!),zero=parseChartDecimal('0'),cross=parseChartDecimal(value.crossesAt!)
 if(compareChartDecimals(minimum,maximum)>=0||compareChartDecimals(minimum,zero)>0||compareChartDecimals(maximum,zero)<0||compareChartDecimals(cross,zero)!==0)throw new RangeError('invalid literal bar scale')
 const count=chart.categories.length,seriesCount=chart.series.length,categoryExtent=column?cx:cy,valueExtent=column?cy:cx
 const categoryReverse=column?category.orientation==='maxMin':category.orientation==='minMax'
 const valueReverse=column?value.orientation==='minMax':value.orientation==='maxMin'
 const denominator=BigInt(2*count*(100*seriesCount+chart.gapWidth)),extent=BigInt(categoryExtent)
 if(200n*extent<denominator)throw new RangeError('literal bar frame cannot retain distinct bars')
 const categoryCoordinate=(numerator:bigint)=>roundRatio((categoryReverse?denominator-numerator:numerator)*extent,denominator)
 const valueCoordinate=(raw:string)=>{let decimal=parseChartDecimal(raw);if(compareChartDecimals(decimal,minimum)<0)decimal=minimum;if(compareChartDecimals(decimal,maximum)>0)decimal=maximum;return chartDecimalCoordinate(decimal,minimum,maximum,valueExtent,valueReverse)}
 const baseline=valueCoordinate('0'),vectors:CartesianBarVector[]=[],indices=new Set<number>()
 chart.series.forEach((series,seriesOrder)=>{
  if(!Number.isInteger(series.index)||series.index<0||series.index>4294967295||indices.has(series.index)||series.order!==seriesOrder||series.values.length!==count||series.colors.length!==count||(series.title!==undefined&&(typeof series.title!=='string'||series.title.length>1024)))throw new RangeError('invalid literal bar series')
  indices.add(series.index)
  series.values.forEach((raw,pointIndex)=>{
   const color=series.colors[pointIndex]!;if(!rgb.test(color))throw new RangeError('invalid literal bar color')
   const numerator=BigInt(2*pointIndex*(100*seriesCount+chart.gapWidth)+chart.gapWidth+200*seriesOrder)
   const a=categoryCoordinate(numerator),b=categoryCoordinate(numerator+200n),end=valueCoordinate(raw)
   const low=Math.min(a,b),high=Math.max(a,b),near=Math.min(baseline,end),far=Math.max(baseline,end)
   const x=column?low:near,y=column?near:low,right=column?high:far,bottom=column?far:high
   vectors.push({seriesIndex:series.index,pointIndex,color,path:[{kind:'moveTo',x,y},{kind:'lineTo',x:right,y},{kind:'lineTo',x:right,y:bottom},{kind:'lineTo',x,y:bottom},{kind:'close'}]})
  })
 })
 const categoryStart=categoryCoordinate(0n)
 if(!category.deleted)vectors.push({axis:'category',stroke:{color:category.color!,widthEmu:category.widthEmu!,cap:'flat',dash:'solid'},path:column?[{kind:'moveTo',x:0,y:baseline},{kind:'lineTo',x:cx,y:baseline}]:[{kind:'moveTo',x:baseline,y:0},{kind:'lineTo',x:baseline,y:cy}]})
 if(!value.deleted)vectors.push({axis:'value',stroke:{color:value.color!,widthEmu:value.widthEmu!,cap:'flat',dash:'solid'},path:column?[{kind:'moveTo',x:categoryStart,y:0},{kind:'lineTo',x:categoryStart,y:cy}]:[{kind:'moveTo',x:0,y:categoryStart},{kind:'lineTo',x:cx,y:categoryStart}]})
 return vectors
}
