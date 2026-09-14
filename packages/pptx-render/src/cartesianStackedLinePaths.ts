import type {NativeLiteralConnected} from '@injoffice/pptx-native'
import {createCartesianConnectedPaths,type CartesianConnectedVector} from './cartesianConnectedPaths.js'
import {createSignedChartStackBands,type SignedStackGrouping} from './chartSignedStacking.js'
import {validateStackedChartGeometry} from './chartStackedGeometryValidation.js'
import {chartRational as r,chartRationalDecimal as decimal,chartRationalSubtract as sub,chartRationalDivide as div,chartRationalCompare as cmp,chartRationalCoordinate as coordinate} from './chartRational.js'
import {clipChartUnitSegment,type ChartRationalPoint} from './chartSegmentClip.js'
import type {RenderPathCommand} from './types.js'

export type CartesianStackedLineInput=Omit<NativeLiteralConnected,'profile'|'dataOrigin'>&{readonly grouping:SignedStackGrouping}
/** Source-ordered algebraic cumulative tops. Disconnected clipped segments never
 * acquire a connector; clipping and percent normalization remain exact until
 * final integer coordinates. Singleton categories have no line segment. */
export function createCartesianStackedLinePaths(chart:CartesianStackedLineInput,cx:number,cy:number):readonly CartesianConnectedVector[]{
 validateStackedChartGeometry(chart,'line')
 const bands=createSignedChartStackBands(chart.series,'line',chart.grouping).bands
 // Keep existing source line axis/style/frame admission unchanged. The scaffold
 // uses zero ordinates only after original decimal data was parsed above.
 const scaffold=createCartesianConnectedPaths({...chart,series:chart.series.map((series,order)=>({...series,order,values:series.values.map(()=> '0')}))},'line',cx,cy)
 const min=decimal(chart.yAxis.min!),span=sub(decimal(chart.yAxis.max!),min),zero=r(0n),one=r(1n)
 const screen=(point:ChartRationalPoint)=>({x:coordinate(point.x,cx,chart.xAxis.orientation==='maxMin'),y:coordinate(point.y,cy,chart.yAxis.orientation==='minMax')})
 const vectors:CartesianConnectedVector[]=[]
 for(const [order,series]of chart.series.entries()){
  const points=bands[order]!.upper.map((value,index)=>({x:r(BigInt(2*index+1),BigInt(2*chart.categories.length)),y:div(sub(value,min),span)}))
  const path:RenderPathCommand[]=[],segmentIndices:number[]=[]
  let previous:ChartRationalPoint|undefined,sourceEnd=false
  for(let i=1;i<points.length;i++){
   const segment=clipChartUnitSegment(points[i-1]!,points[i]!)
   if(!segment){previous=undefined;sourceEnd=false;continue}
   const continuous=previous!==undefined&&sourceEnd&&cmp(segment.startParameter,zero)===0&&cmp(previous.x,segment.start.x)===0&&cmp(previous.y,segment.start.y)===0
   if(!continuous)path.push({kind:'moveTo',...screen(segment.start)})
   path.push({kind:'lineTo',...screen(segment.end)})
   segmentIndices.push(i-1);previous=segment.end;sourceEnd=cmp(segment.endParameter,one)===0
  }
  if(path.length>510)throw new RangeError('stacked line path budget exceeded')
  vectors.push({seriesIndex:series.index,segmentIndices,path,stroke:{color:series.color,widthEmu:series.widthEmu,cap:'flat',join:'round',dash:'solid'}})
 }
 return [...vectors,...scaffold.filter(vector=>vector.axis!==undefined)]
}
