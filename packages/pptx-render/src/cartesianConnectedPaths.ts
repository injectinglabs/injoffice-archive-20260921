import type {NativeLiteralBarAxis,NativeLiteralConnected} from '@injoffice/pptx-native'
import type {RenderPathCommand,RenderStroke} from './types.js'
import {chartRational as r,chartRationalDecimal as decimal,chartRationalSubtract as subtract,chartRationalDivide as divide,chartRationalCompare as compare,chartRationalCoordinate as coordinate} from './chartRational.js'
import {clipChartUnitSegment,type ChartRationalPoint} from './chartSegmentClip.js'

export interface CartesianConnectedVector {readonly seriesIndex?:number;readonly segmentIndices?:readonly number[];readonly axis?:'x'|'y';readonly path:readonly RenderPathCommand[];readonly stroke:RenderStroke}
const zero=r(0n),one=r(1n),rgb=/^#[0-9A-F]{6}$/
function exactScale(axis:NativeLiteralBarAxis){
 const min=decimal(axis.min!),max=decimal(axis.max!),cross=decimal(axis.crossesAt!)
 if(compare(min,max)>=0||compare(min,zero)>0||compare(max,zero)<0||compare(cross,zero)!==0)throw new RangeError('invalid explicit connected chart scale')
 const span=subtract(max,min)
 return (value:string)=>divide(subtract(decimal(value),min),span)
}
function validateAxis(axis:NativeLiteralBarAxis):void{
 if(!Number.isInteger(axis.id)||axis.id<0||axis.id>4294967295||!Number.isInteger(axis.crossAxisId)||axis.crossAxisId<0||axis.crossAxisId>4294967295||!['minMax','maxMin'].includes(axis.orientation)||typeof axis.deleted!=='boolean'||(axis.deleted?(axis.color!==undefined||axis.widthEmu!==undefined):(!rgb.test(axis.color??'')||!Number.isInteger(axis.widthEmu)||axis.widthEmu!<1||axis.widthEmu!>20116800)))throw new RangeError('invalid connected chart axis')
}
function same(a:ChartRationalPoint,b:ChartRationalPoint):boolean{return compare(a.x,b.x)===0&&compare(a.y,b.y)===0}

/** Shared implementation; source family wrappers require their exact profile. */
export function createCartesianConnectedPaths(chart:Omit<NativeLiteralConnected,'profile'|'dataOrigin'>,family:'line'|'scatter',cx:number,cy:number):readonly CartesianConnectedVector[]{
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>281474976710655||cy>281474976710655||!['line','scatter'].includes(family)||chart.series.length<1||chart.series.length>16)throw new RangeError('invalid connected chart frame/profile')
 const scatter=family==='scatter',xAxis=chart.xAxis,yAxis=chart.yAxis
 validateAxis(xAxis);validateAxis(yAxis)
 if(xAxis.id===yAxis.id||xAxis.crossAxisId!==yAxis.id||yAxis.crossAxisId!==xAxis.id||xAxis.position!=='b'||yAxis.position!=='l')throw new RangeError('invalid connected chart axes')
 if(scatter?chart.categories.length!==0:(chart.categories.length<1||chart.categories.length>256||chart.categories.some(c=>typeof c!=='string')||chart.categories.reduce((n,c)=>n+c.length,0)>32768||xAxis.min!==undefined||xAxis.max!==undefined||xAxis.crossesAt!==undefined))throw new RangeError('invalid connected chart categories')
 const yScale=exactScale(yAxis),xScale=scatter?exactScale(xAxis):undefined
 const vectors:CartesianConnectedVector[]=[],seen=new Set<number>()
 const screen=(point:ChartRationalPoint)=>({x:coordinate(point.x,cx,xAxis.orientation==='maxMin'),y:coordinate(point.y,cy,yAxis.orientation==='minMax')})
 chart.series.forEach((series,order)=>{
  if(!Number.isInteger(series.index)||series.index<0||series.index>4294967295||seen.has(series.index)||series.order!==order||series.values.length<1||series.values.length>256||!rgb.test(series.color)||!Number.isInteger(series.widthEmu)||series.widthEmu<1||series.widthEmu>20116800||(series.title!==undefined&&(typeof series.title!=='string'||series.title.length>1024)))throw new RangeError('invalid connected chart series')
  seen.add(series.index)
  if(scatter?(!series.xValues||series.xValues.length!==series.values.length):(series.xValues!==undefined||series.values.length!==chart.categories.length))throw new RangeError('connected chart point counts differ')
  const points=series.values.map((value,index):ChartRationalPoint=>({x:scatter?xScale!(series.xValues![index]!):r(BigInt(2*index+1),BigInt(2*chart.categories.length)),y:yScale(value)}))
  const path:RenderPathCommand[]=[],segmentIndices:number[]=[]
  let previous:ChartRationalPoint|undefined,previousEndedAtSource=false
  for(let i=1;i<points.length;i++){
   const segment=clipChartUnitSegment(points[i-1]!,points[i]!)
   if(!segment){previous=undefined;previousEndedAtSource=false;continue}
   const continuous=previous!==undefined&&previousEndedAtSource&&compare(segment.startParameter,zero)===0&&same(previous,segment.start)
   if(!continuous)path.push({kind:'moveTo',...screen(segment.start)})
   path.push({kind:'lineTo',...screen(segment.end)})
   segmentIndices.push(i-1);previous=segment.end;previousEndedAtSource=compare(segment.endParameter,one)===0
  }
  // <=255 segments, each requiring at most a move and line:510 commands.
  if(path.length>510)throw new RangeError('connected chart path budget exceeded')
  vectors.push({seriesIndex:series.index,segmentIndices,path,stroke:{color:series.color,widthEmu:series.widthEmu,cap:'flat',join:'round',dash:'solid'}})
 })
 const xCross=scatter?xScale!('0'):zero,yCross=yScale('0')
 if(!xAxis.deleted){const y=coordinate(yCross,cy,yAxis.orientation==='minMax');vectors.push({axis:'x',path:[{kind:'moveTo',x:0,y},{kind:'lineTo',x:cx,y}],stroke:{color:xAxis.color!,widthEmu:xAxis.widthEmu!,cap:'flat',dash:'solid'}})}
 if(!yAxis.deleted){const x=coordinate(xCross,cx,xAxis.orientation==='maxMin');vectors.push({axis:'y',path:[{kind:'moveTo',x,y:0},{kind:'lineTo',x,y:cy}],stroke:{color:yAxis.color!,widthEmu:yAxis.widthEmu!,cap:'flat',dash:'solid'}})}
 return vectors
}
