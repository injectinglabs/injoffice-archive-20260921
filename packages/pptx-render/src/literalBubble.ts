import {validNativeLiteralBubble,type NativeLiteralBubble} from '@injoffice/pptx-native'
import {createChartBubbleGeometry} from './chartBubbleGeometry.js'
import {chartRational as r,chartRationalDecimal as decimal,chartRationalSubtract as sub,chartRationalDivide as div,chartRationalCoordinate as coordinate} from './chartRational.js'
import type {RenderPathCommand,RenderStroke,RenderRect} from './types.js'

export interface LiteralBubbleVector {
 readonly seriesIndex?:number;readonly pointIndex?:number;readonly axis?:'x'|'y'
 readonly path:readonly RenderPathCommand[];readonly color?:string;readonly stroke?:RenderStroke
 /** Complete circle hull, including portions outside the plot. */
 readonly bubbleBounds?:RenderRect
}
export const BUBBLE_PREVIEW_POLICY='plot-minor-radius-v1'
export const BUBBLE_PREVIEW_DISCLOSURE='Source bubble sizes use plot-minor-radius-v1: the global maximum radius is one tenth of the smaller host plot extent at scale 100; area/width and source scale are retained. Circles paint in XML series sequence then point index and clip to the whole plot. This is host sizing, not PowerPoint plot layout.'

/** Shared visual-value implementation, called only after the appropriate literal
 * or source-bound workbook admission. The temporary validation shape below is
 * never published as a chart or admitted as source data. */
export function createCartesianBubblePaths(chart:Omit<NativeLiteralBubble,'profile'|'dataOrigin'>,cx:number,cy:number):readonly LiteralBubbleVector[]{
 const {bubbleScale,sizeRepresents,series,xAxis:x,yAxis:y}=chart
 if(!validNativeLiteralBubble({profile:'literal-bubble-v1',dataOrigin:'literal',bubbleScale,sizeRepresents,series,xAxis:x,yAxis:y}))throw new RangeError('invalid bubble visual values')
 const geometry=createChartBubbleGeometry(series,{xMin:x.min!,xMax:x.max!,yMin:y.min!,yMax:y.max!,reverseX:x.orientation==='maxMin',reverseY:y.orientation==='maxMin',bubbleScale,sizeRepresents,sizingPolicy:BUBBLE_PREVIEW_POLICY},cx,cy)
 const paints=new Map(series.map(s=>[s.index,s.colors]))
 const vectors:LiteralBubbleVector[]=geometry.map(g=>({seriesIndex:g.seriesIndex,pointIndex:g.pointIndex,path:g.path,color:paints.get(g.seriesIndex)![g.pointIndex]!,bubbleBounds:{x:g.centerX-g.radius,y:g.centerY-g.radius,cx:2*g.radius,cy:2*g.radius}}))
 const cross=(min:string,max:string,extent:number,reverse:boolean)=>coordinate(div(sub(r(0n),decimal(min)),sub(decimal(max),decimal(min))),extent,reverse)
 if(!x.deleted){const at=cross(y.min!,y.max!,cy,y.orientation==='minMax');vectors.push({axis:'x',path:[{kind:'moveTo',x:0,y:at},{kind:'lineTo',x:cx,y:at}],stroke:{color:x.color!,widthEmu:x.widthEmu!,cap:'flat',dash:'solid'}})}
 if(!y.deleted){const at=cross(x.min!,x.max!,cx,x.orientation==='maxMin');vectors.push({axis:'y',path:[{kind:'moveTo',x:at,y:0},{kind:'lineTo',x:at,y:cy}],stroke:{color:y.color!,widthEmu:y.widthEmu!,cap:'flat',dash:'solid'}})}
 return vectors
}
export function createNativeLiteralBubblePaths(chart:NativeLiteralBubble,cx:number,cy:number):readonly LiteralBubbleVector[]{
 if(!validNativeLiteralBubble(chart))throw new RangeError('expected qualified literal bubble profile')
 return createCartesianBubblePaths(chart,cx,cy)
}
