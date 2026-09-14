import {chartRational as r,chartRationalDecimal as decimal,chartRationalAdd as add,chartRationalSubtract as sub,chartRationalMultiply as mul,chartRationalDivide as div,chartRationalCompare as compare,type ChartRational} from './chartRational.js'
import {chartRoundedSquareRoot} from './chartBoundedSqrt.js'
import type {RenderPathCommand} from './types.js'
export interface ChartBubbleSeries {readonly index:number;readonly order:number;readonly xValues:readonly string[];readonly values:readonly string[];readonly sizes:readonly string[]}
export interface ChartBubbleScale {readonly xMin:string;readonly xMax:string;readonly yMin:string;readonly yMax:string;readonly reverseX?:boolean;readonly reverseY?:boolean;readonly bubbleScale:number;readonly sizeRepresents:'area'|'w';readonly sizingPolicy:'plot-minor-radius-v1'}
export interface ChartBubbleGeometry {readonly seriesIndex:number;readonly pointIndex:number;readonly centerX:number;readonly centerY:number;readonly radius:number;readonly path:readonly RenderPathCommand[]}
const zero=r(0n),one=r(1n)
function integer(value:number,low:number,high:number):boolean{return Number.isSafeInteger(value)&&!Object.is(value,-0)&&value>=low&&value<=high}
function round(value:ChartRational):number {
 const sign=value.numerator<0n?-1n:1n,n=value.numerator*sign
 const rounded=sign*((2n*n+value.denominator)/(2n*value.denominator))
 if(rounded>BigInt(Number.MAX_SAFE_INTEGER)||rounded< -BigInt(Number.MAX_SAFE_INTEGER))throw new RangeError('bubble coordinate exceeds integer bounds')
 return Number(rounded)
}
/** Dedicated exact source-value geometry. The caller must clip the ENTIRE
 * circles to the plot rectangle, not discard off-plot centers or clip centers.
 * plot-minor-radius-v1 is a disclosed host sizing policy: global maximum radius
 * is 1/10 of the smaller plot extent at source scale100. It is not Office sizing.
 * Radius and center quantize once to nearest EMU; positive sub-EMU bubbles may
 * disappear. Source paint follows ascending series order then point index. */
export function createChartBubbleGeometry(series:readonly ChartBubbleSeries[],scale:ChartBubbleScale,cx:number,cy:number):readonly ChartBubbleGeometry[]{
 if(!integer(cx,1,281474976710655)||!integer(cy,1,281474976710655)||!scale||scale.sizingPolicy!=='plot-minor-radius-v1'||!integer(scale.bubbleScale,0,300)||!['area','w'].includes(scale.sizeRepresents)||(scale.reverseX!==undefined&&typeof scale.reverseX!=='boolean')||(scale.reverseY!==undefined&&typeof scale.reverseY!=='boolean')||!Array.isArray(series)||series.length<1||series.length>16)throw new RangeError('invalid bubble profile/frame')
 const xMin=decimal(scale.xMin),xMax=decimal(scale.xMax),yMin=decimal(scale.yMin),yMax=decimal(scale.yMax)
 if(compare(xMin,xMax)>=0||compare(yMin,yMax)>=0)throw new RangeError('invalid bubble scale')
 const indices=new Set<number>(),source:{s:ChartBubbleSeries;x:ChartRational[];y:ChartRational[];z:ChartRational[]}[]=[]
 let maximum=zero
 for(let order=0;order<series.length;order++){
  const s=series[order]!
  if(!s||!integer(s.index,0,4294967295)||indices.has(s.index)||!integer(s.order,0,15)||s.order!==order||!Array.isArray(s.values)||!Array.isArray(s.xValues)||!Array.isArray(s.sizes)||s.values.length<1||s.values.length>256||s.xValues.length!==s.values.length||s.sizes.length!==s.values.length)throw new RangeError('invalid bubble series')
  indices.add(s.index)
  const x=s.xValues.map(decimal),y=s.values.map(decimal),z=s.sizes.map(decimal)
  // for-of also detects sparse arrays (map alone would retain the holes).
  for(const collection of [x,y,z])for(const value of collection)if(!value)throw new RangeError('sparse bubble values')
  for(const size of z){if(compare(size,zero)<0)throw new RangeError('negative bubble size outside profile');if(compare(size,maximum)>0)maximum=size}
  source.push({s,x,y,z})
 }
 if(maximum.numerator===0n||scale.bubbleScale===0)return []
 const maxRadius=r(BigInt(Math.min(cx,cy))*BigInt(scale.bubbleScale),1000n),result:ChartBubbleGeometry[]=[]
 for(const {s,x,y,z} of source)for(let point=0;point<z.length;point++){
  if(z[point]!.numerator===0n)continue
  const ratio=div(z[point]!,maximum),radius=scale.sizeRepresents==='w'?round(mul(maxRadius,ratio)):chartRoundedSquareRoot(mul(mul(maxRadius,maxRadius),ratio))
  if(radius===0)continue
  let unitX=div(sub(x[point]!,xMin),sub(xMax,xMin)),unitY=div(sub(y[point]!,yMin),sub(yMax,yMin))
  if(scale.reverseX)unitX=sub(one,unitX)
  if(!scale.reverseY)unitY=sub(one,unitY)
  const centerX=mul(unitX,r(BigInt(cx))),centerY=mul(unitY,r(BigInt(cy)))
  // Conservative rectangle envelope includes half-EMU center rounding. Exact
  // comparisons cull arbitrarily distant source values before Number conversion.
  const margin=r(BigInt(2*radius+1),2n)
  if(compare(add(centerX,margin),zero)<0||compare(sub(centerX,margin),r(BigInt(cx)))>0||compare(add(centerY,margin),zero)<0||compare(sub(centerY,margin),r(BigInt(cy)))>0)continue
  const px=round(centerX),py=round(centerY)
  const path:RenderPathCommand[]=[{kind:'moveTo',x:px+radius,y:py},{kind:'arcTo',x:px-radius,y:py,rx:radius,ry:radius,largeArc:false,clockwise:true},{kind:'arcTo',x:px+radius,y:py,rx:radius,ry:radius,largeArc:false,clockwise:true},{kind:'close'}]
  result.push({seriesIndex:s.index,pointIndex:point,centerX:px,centerY:py,radius,path})
 }
 return result
}
