import type {NativeLiteralRadar} from '@injoffice/pptx-native'
import {validNativeLiteralRadar} from '@injoffice/pptx-native'
import {chartRational as r,chartRationalDecimal as decimal,chartRationalAdd as add,chartRationalSubtract as sub,chartRationalMultiply as mul,chartRationalDivide as div,chartRationalCoordinate as coordinate} from './chartRational.js'
import type {RenderPathCommand,RenderRect,RenderStroke} from './types.js'

export const RADAR_PREVIEW_POLICY='source-radial-plot-v1'
export interface RadarSeriesVector {
 readonly seriesIndex?:number
 readonly sourceOrder?:number
 readonly axis?:'value'
 readonly path:readonly RenderPathCommand[]
 readonly stroke:RenderStroke
 readonly color?:string
 /** Includes stroke and a separate one-EMU numerical envelope. */
 readonly inkBounds:RenderRect
}
/** Dedicated geometry only: caller owns source admission, axis paint, plot clip
 * and complete transformed ink qualification. No source array is rewritten. */
export function createChartRadarSeriesGeometry(chart:NativeLiteralRadar,cx:number,cy:number):readonly RadarSeriesVector[]{
 if(!validNativeLiteralRadar(chart))throw new RangeError('unqualified radar chart profile')
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>100_000_000||cy>100_000_000)throw new RangeError('radar frame coordinate budget exceeded')
 const count=chart.categories.length,radius=r(BigInt(Math.min(cx,cy)),2n),minimum=decimal(chart.valueAxis.min!),span=sub(decimal(chart.valueAxis.max!),minimum)
 const directions=radarDirections(count,chart.categoryAxis.orientation==='maxMin',cx,cy)
 return chart.series.map(series=>{
  const points=series.values.map((raw,index)=>{
   let radial=div(sub(decimal(raw),minimum),span)
   if(chart.valueAxis.orientation==='maxMin')radial=sub(r(1n),radial)
   return radarPoint(radial,directions[index]!,radius,cx,cy)
  })
  const path:RenderPathCommand[]=points.map((point,index)=>({kind:index===0?'moveTo':'lineTo',...point}));path.push({kind:'close'})
  if(path.length>257)throw new RangeError('radar path command budget exceeded')
  return {seriesIndex:series.index,sourceOrder:series.order,path,stroke:{color:series.color,widthEmu:series.widthEmu,cap:'flat',join:'round',dash:'solid'},...(series.fill?{color:series.fill}:{}),inkBounds:radarInkBounds(points,series.widthEmu)}
 })
}
function radarDirections(count:number,reverse:boolean,cx:number,cy:number){
 return Array.from({length:count},(_,index)=>{
  const quarter=index*4,countQuarter=quarter/count,cardinal=quarter%count===0
  const angle=index*2*Math.PI/count
  // Same explicitly non-formal libm qualification policy as sourceAffine.
  // Scope is finite [0,2pi), n<=256 and radius<=50 million EMU.
  const allowance=cardinal?0:64*Number.EPSILON*(Math.abs(angle)+1)
  if(allowance*Math.min(cx,cy)/2>0.00001)throw new RangeError('radar angular uncertainty budget exceeded')
  const sin=cardinal?[0,1,0,-1][countQuarter]!:Math.sin(angle),cos=cardinal?[1,0,-1,0][countQuarter]!:Math.cos(angle)
  if(!Number.isFinite(sin)||!Number.isFinite(cos)||Math.abs(sin)>1||Math.abs(cos)>1)throw new RangeError('invalid radar angular projection')
  return {sin:decimal(String(reverse?-sin:sin)),cos:decimal(String(cos))}
 })
}
function radarPoint(radial:ReturnType<typeof r>,direction:ReturnType<typeof radarDirections>[number],radius:ReturnType<typeof r>,cx:number,cy:number){
 const half=r(1n,2n),length=mul(radius,radial)
 const x=add(half,div(mul(length,direction.sin),r(BigInt(cx)))),y=sub(half,div(mul(length,direction.cos),r(BigInt(cy))))
 return {x:coordinate(x,cx),y:coordinate(y,cy)}
}
function radarInkBounds(points:readonly {x:number;y:number}[],width:number):RenderRect{
 const pad=Math.ceil(width/2)+1
 let left=Infinity,top=Infinity,right=-Infinity,bottom=-Infinity
 for(const point of points){left=Math.min(left,point.x);top=Math.min(top,point.y);right=Math.max(right,point.x);bottom=Math.max(bottom,point.y)}
 return {x:left-pad,y:top-pad,cx:right-left+2*pad,cy:bottom-top+2*pad}
}
/** Reference-qualified no-label value-axis spokes. The category axis supplies
 * angular order, not a separate visible outline in this bounded Office profile.
 * Standard spokes precede data; filled spokes follow data. */
export function createChartRadarGeometry(chart:NativeLiteralRadar,cx:number,cy:number):readonly RadarSeriesVector[]{
 const data=createChartRadarSeriesGeometry(chart,cx,cy)
 if(chart.valueAxis.deleted)return data
 const points=radarDirections(chart.categories.length,chart.categoryAxis.orientation==='maxMin',cx,cy).map(d=>radarPoint(r(1n),d,r(BigInt(Math.min(cx,cy)),2n),cx,cy)),center={x:Math.round(cx/2),y:Math.round(cy/2)}
 const path:RenderPathCommand[]=points.flatMap(point=>[{kind:'moveTo' as const,...center},{kind:'lineTo' as const,...point}])
 if(path.length>512)throw new RangeError('radar spoke path command budget exceeded')
 const axis:RadarSeriesVector={axis:'value',path,stroke:{color:chart.valueAxis.color!,widthEmu:chart.valueAxis.widthEmu!,cap:'flat',join:'round',dash:'solid'},inkBounds:radarInkBounds([...points,center],chart.valueAxis.widthEmu!)}
 return chart.style==='filled'?[...data,axis]:[axis,...data]
}
