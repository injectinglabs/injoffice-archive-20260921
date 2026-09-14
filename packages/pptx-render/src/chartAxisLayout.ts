import type {NativeLiteralBarAxis} from '@injoffice/pptx-native'
import type {RenderRect,RenderTextBodyNode} from './types.js'
import {createChartAxisTicks} from './chartAxisTicks.js'
import {chartRational,chartRationalCoordinate,chartRationalDecimal,type ChartRational} from './chartRational.js'

export const CHART_AXIS_LAYOUT_POLICY='supplied-outline-margins-v1' as const
export const CHART_AXIS_GAP_EMU=12700
export const CHART_AXIS_TICK_EMU=38100
export interface ChartAxisLabelInput {axis:NativeLiteralBarAxis;perpendicular:NativeLiteralBarAxis;horizontal:boolean;/** Area source minimum crossing; absent keeps existing axis policy. */crossingAtMinimum?:true;categories?:readonly string[]}
interface MeasuredLabel {input:ChartAxisLabelInput;text:string;position:ChartRational;body:RenderTextBodyNode;bounds:RenderRect;side:'left'|'right'|'top'|'bottom'}
export interface PlacedChartAxisLabel {body:RenderTextBodyNode;bounds:RenderRect;x:number;y:number;text:string}
export interface ChartAxisLayout {plot:RenderRect;labels:readonly PlacedChartAxisLabel[]}

/** Host layout, not Office plot-layout reproduction. The source font, text,
 * tick positions and perpendicular reversal survive; no label is dropped. */
export async function layoutChartAxes(inputs:readonly ChartAxisLabelInput[],cx:number,cy:number,measure:(text:string,axis:NativeLiteralBarAxis,index:number)=>Promise<{body:RenderTextBodyNode;bounds:RenderRect}>):Promise<ChartAxisLayout>{
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>281474976710655||cy>281474976710655)throw new RangeError('invalid chart axis frame')
 const measured:MeasuredLabel[]=[]
 let left=0,right=0,top=0,bottom=0
 for(const input of inputs){
  const labels=input.axis.labels;if(!labels)continue
  const halfStroke=Math.ceil(input.axis.widthEmu!/2)
  if(!Number.isSafeInteger(halfStroke)||halfStroke<1)throw new RangeError('invalid source axis stroke')
  if(labels.majorTickMark==='out'){const reserve=CHART_AXIS_TICK_EMU+halfStroke;left=Math.max(left,reserve);right=Math.max(right,reserve);top=Math.max(top,reserve);bottom=Math.max(bottom,reserve)}
  // low/high refer to the perpendicular value order, not physical screen sides.
  const low=labels.position==='low',normal=input.perpendicular.orientation==='minMax'
  const side=input.horizontal?(low===normal?'bottom':'top'):(low===normal?'left':'right')
  const ticks=input.categories?input.categories.map((label,i)=>({label,position:chartRational(BigInt(2*i+1),BigInt(2*input.categories!.length))})):createChartAxisTicks(input.axis.min!,input.axis.max!,labels.majorUnit!,labels.numberFormat!)
  for(const tick of ticks){
   if(measured.length>=512)throw new RangeError('chart axis label count exceeded')
   const result=await measure(tick.label,input.axis,measured.length),b=result.bounds
   if([b.x,b.y,b.cx,b.cy].some(v=>!Number.isSafeInteger(v))||b.cx<=0||b.cy<=0||b.cx>cx||b.cy>cy)throw new RangeError('chart axis label cannot fit the source frame')
   measured.push({...result,input,text:tick.label,position:tick.position,side})
   const gap=CHART_AXIS_GAP_EMU+CHART_AXIS_TICK_EMU+Math.ceil(input.axis.widthEmu!/2)
   if(input.horizontal){left=Math.max(left,Math.ceil(b.cx/2));right=Math.max(right,Math.ceil(b.cx/2));if(side==='top')top=Math.max(top,b.cy+gap);else bottom=Math.max(bottom,b.cy+gap)}
   else {top=Math.max(top,Math.ceil(b.cy/2));bottom=Math.max(bottom,Math.ceil(b.cy/2));if(side==='left')left=Math.max(left,b.cx+gap);else right=Math.max(right,b.cx+gap)}
  }
 }
 // Endpoint labels on perpendicular axes may meet at a corner. Their
 // measured half-extents set clearance before the plot frame is finalized.
 const gaps=new Map<ChartAxisLabelInput,number>()
 for(const input of inputs){
  const own=measured.filter(label=>label.input===input);if(!own.length)continue
  let crossExtent=0
  for(const other of measured)if(other.input.horizontal!==input.horizontal)crossExtent=Math.max(crossExtent,input.horizontal?other.bounds.cy:other.bounds.cx)
  const gap=Math.max(CHART_AXIS_GAP_EMU+CHART_AXIS_TICK_EMU+Math.ceil(input.axis.widthEmu!/2),Math.ceil(crossExtent/2)+CHART_AXIS_GAP_EMU)
  gaps.set(input,gap)
  for(const label of own){const b=label.bounds;if(label.side==='left')left=Math.max(left,b.cx+gap);if(label.side==='right')right=Math.max(right,b.cx+gap);if(label.side==='top')top=Math.max(top,b.cy+gap);if(label.side==='bottom')bottom=Math.max(bottom,b.cy+gap)}
 }
 const plot={x:left,y:top,cx:cx-left-right,cy:cy-top-bottom}
 if(plot.cx<1||plot.cy<1)throw new RangeError('axis margins leave no bounded plot frame')
 const placed:PlacedChartAxisLabel[]=[]
 for(const label of measured){
  const {input,bounds:b,side}=label
  const coordinate=chartRationalCoordinate(label.position,input.horizontal?plot.cx:plot.cy,input.horizontal?input.axis.orientation==='maxMin':input.axis.orientation==='minMax')
  const gap=gaps.get(input)!
  const x=input.horizontal?plot.x+coordinate-Math.floor(b.cx/2):side==='left'?plot.x-gap-b.cx:plot.x+plot.cx+gap
  const y=input.horizontal?(side==='top'?plot.y-gap-b.cy:plot.y+plot.cy+gap):plot.y+coordinate-Math.floor(b.cy/2)
  const rect={x,y,cx:b.cx,cy:b.cy}
  if(x<0||y<0||x+b.cx>cx||y+b.cy>cy)throw new RangeError('axis label exceeds chart frame')
  if(placed.some(p=>x<p.bounds.x+p.bounds.cx&&x+b.cx>p.bounds.x&&y<p.bounds.y+p.bounds.cy&&y+b.cy>p.bounds.y))throw new RangeError('source axis labels overlap under supplied-font layout')
  placed.push({body:label.body,bounds:rect,x:x-b.x,y:y-b.y,text:label.text})
 }
 return {plot,labels:placed}
}

/** Explicit out ticks qualify only at a plot edge; stroke widths remain source values. */
export function chartAxisTickVectors(inputs:readonly ChartAxisLabelInput[],plot:RenderRect):readonly {path:readonly import('./types.js').RenderPathCommand[];stroke:import('./types.js').RenderStroke}[]{
 const result:{path:import('./types.js').RenderPathCommand[];stroke:import('./types.js').RenderStroke}[]=[]
 for(const input of inputs){
  const {axis,perpendicular,horizontal,categories}=input,labels=axis.labels
  if(!labels||labels.majorTickMark!=='out')continue
  const low=input.crossingAtMinimum===true||perpendicular.min===undefined||chartRationalDecimal(perpendicular.min).numerator===0n
  const normal=perpendicular.orientation==='minMax'
  const near=horizontal?low!==normal:low===normal
  const edge=horizontal?(near?plot.y:plot.y+plot.cy):(near?plot.x:plot.x+plot.cx)
  const positions=categories?categories.map((_,i)=>chartRational(BigInt(2*i+1),BigInt(2*categories.length))):createChartAxisTicks(axis.min!,axis.max!,labels.majorUnit!,labels.numberFormat!).map(t=>t.position)
  for(const position of positions){
   const p=chartRationalCoordinate(position,horizontal?plot.cx:plot.cy,horizontal?axis.orientation==='maxMin':axis.orientation==='minMax')+(horizontal?plot.x:plot.y)
   const end=edge+(near?-CHART_AXIS_TICK_EMU:CHART_AXIS_TICK_EMU)
   result.push({path:horizontal?[{kind:'moveTo',x:p,y:edge},{kind:'lineTo',x:p,y:end}]:[{kind:'moveTo',x:edge,y:p},{kind:'lineTo',x:end,y:p}],stroke:{color:axis.color!,widthEmu:axis.widthEmu!,cap:'flat',dash:'solid'}})
  }
 }
 return result
}
