import type {NativeArrowEnd} from '@injoffice/pptx-native'
import type {RenderPathCommand} from '@injoffice/pptx-render'
import {previewArrowShaftInset} from './arrows.js'

type Point={readonly x:number;readonly y:number}
type Segment=
 |{readonly kind:'line';readonly from:Point;readonly to:Point}
 |{readonly kind:'quad';readonly from:Point;readonly c1:Point;readonly to:Point}
 |{readonly kind:'cubic';readonly from:Point;readonly c1:Point;readonly c2:Point;readonly to:Point}
export interface ConnectorShaftPreview {
 /** SVG path data of the shaft after arrow-v1 endpoint insets. */
 readonly d:string
 readonly head:{readonly tip:Point;readonly direction:Point}
 readonly tail:{readonly tip:Point;readonly direction:Point}
 /** True for the pre-existing exact straight two-command connector. */
 readonly straight:boolean
}

const SPLIT_ITERATIONS=48
const lerp=(a:Point,b:Point,t:number):Point=>({x:a.x+(b.x-a.x)*t,y:a.y+(b.y-a.y)*t})
const distance=(a:Point,b:Point)=>Math.hypot(b.x-a.x,b.y-a.y)
const same=(a:Point,b:Point)=>a.x===b.x&&a.y===b.y
const finitePoint=(p:Point)=>Number.isFinite(p.x)&&Number.isFinite(p.y)

/** Evaluates a segment at parameter t (Bernstein form). */
function at(segment:Segment,t:number):Point {
 switch(segment.kind){
  case 'line':return lerp(segment.from,segment.to,t)
  case 'quad':return lerp(lerp(segment.from,segment.c1,t),lerp(segment.c1,segment.to,t),t)
  case 'cubic':{const q0=lerp(segment.from,segment.c1,t),q1=lerp(segment.c1,segment.c2,t),q2=lerp(segment.c2,segment.to,t);return lerp(lerp(q0,q1,t),lerp(q1,q2,t),t)}
 }
}

/** de Casteljau split; returns the [t,1] portion. */
function tailPortion(segment:Segment,t:number):Segment {
 switch(segment.kind){
  case 'line':return {kind:'line',from:lerp(segment.from,segment.to,t),to:segment.to}
  case 'quad':{const q0=lerp(segment.from,segment.c1,t),q1=lerp(segment.c1,segment.to,t);return {kind:'quad',from:lerp(q0,q1,t),c1:q1,to:segment.to}}
  case 'cubic':{const q0=lerp(segment.from,segment.c1,t),q1=lerp(segment.c1,segment.c2,t),q2=lerp(segment.c2,segment.to,t),r0=lerp(q0,q1,t),r1=lerp(q1,q2,t);return {kind:'cubic',from:lerp(r0,r1,t),c1:r1,c2:q2,to:segment.to}}
 }
}

function reverse(segment:Segment):Segment {
 switch(segment.kind){
  case 'line':return {kind:'line',from:segment.to,to:segment.from}
  case 'quad':return {kind:'quad',from:segment.to,c1:segment.c1,to:segment.from}
  case 'cubic':return {kind:'cubic',from:segment.to,c1:segment.c2,c2:segment.c1,to:segment.from}
 }
}

/** Removes `inset` EMU measured as straight-line distance from the segment
 * start so a filled arrowhead covers the shaft end instead of a flat stub.
 * Straight segments trim exactly; curves use a bounded bisection on the
 * chord distance and a de Casteljau split. Refuses when the segment is too
 * short instead of drawing a reversed or empty shaft. */
function trimStart(segment:Segment,inset:number):Segment {
 if(inset<=0)return segment
 if(segment.kind==='line'){
  const length=distance(segment.from,segment.to)
  if(length<=inset)throw new Error('Source connector is too short for arrow-v1 endpoint geometry')
  return tailPortion(segment,inset/length)
 }
 if(distance(segment.from,segment.to)<=inset)throw new Error('Source connector is too short for arrow-v1 endpoint geometry')
 let lo=0,hi=1
 for(let i=0;i<SPLIT_ITERATIONS;i++){const mid=(lo+hi)/2;if(distance(at(segment,mid),segment.from)<inset)lo=mid;else hi=mid}
 return tailPortion(segment,hi)
}

function firstTangent(segment:Segment):Point {
 if(segment.kind==='line')return segment.to
 if(segment.kind==='quad')return same(segment.c1,segment.from)?segment.to:segment.c1
 return !same(segment.c1,segment.from)?segment.c1:!same(segment.c2,segment.from)?segment.c2:segment.to
}

function segmentPath(segment:Segment):string {
 switch(segment.kind){
  case 'line':return `L${segment.to.x} ${segment.to.y}`
  case 'quad':return `Q${segment.c1.x} ${segment.c1.y} ${segment.to.x} ${segment.to.y}`
  case 'cubic':return `C${segment.c1.x} ${segment.c1.y} ${segment.c2.x} ${segment.c2.y} ${segment.to.x} ${segment.to.y}`
 }
}

/** Builds the arrow-v1 shaft and endpoint tangents for an open connector path
 * (exact straight line or an evaluated connector-preset polyline/curve).
 * Explicit InjOffice preview geometry, not PowerPoint-equivalent metrics. */
export function previewConnectorShaft(path:readonly RenderPathCommand[],headEnd:Readonly<NativeArrowEnd>|undefined,tailEnd:Readonly<NativeArrowEnd>|undefined,strokeWidth:number):ConnectorShaftPreview {
 const start=path[0]
 if(path.length<2||start?.kind!=='moveTo'||!finitePoint(start))throw new Error('Arrow shaft must be an open source connector path')
 const segments:Segment[]=[]
 let pen:Point={x:start.x,y:start.y}
 for(const command of path.slice(1)){
  let segment:Segment
  switch(command.kind){
   case 'lineTo':segment={kind:'line',from:pen,to:{x:command.x,y:command.y}};break
   case 'quadBezierTo':segment={kind:'quad',from:pen,c1:{x:command.x1,y:command.y1},to:{x:command.x,y:command.y}};break
   case 'cubicBezierTo':segment={kind:'cubic',from:pen,c1:{x:command.x1,y:command.y1},c2:{x:command.x2,y:command.y2},to:{x:command.x,y:command.y}};break
   default:throw new Error('Arrow shaft must be an open source connector path')
  }
  if(!finitePoint(segment.to)||(segment.kind!=='line'&&!finitePoint(segment.c1))||(segment.kind==='cubic'&&!finitePoint(segment.c2)))throw new Error('Arrow shaft must be an open source connector path')
  segments.push(segment);pen=segment.to
 }
 const first=segments[0]!,last=segments[segments.length-1]!
 const head={tip:first.from,direction:{x:first.from.x-firstTangent(first).x,y:first.from.y-firstTangent(first).y}}
 const reversedLast=reverse(last),tailTangent=firstTangent(reversedLast)
 const tail={tip:last.to,direction:{x:last.to.x-tailTangent.x,y:last.to.y-tailTangent.y}}
 const headInset=previewArrowShaftInset(headEnd,strokeWidth),tailInset=previewArrowShaftInset(tailEnd,strokeWidth)
 const trimmed=[...segments]
 trimmed[0]=trimStart(trimmed[0]!,headInset)
 trimmed[trimmed.length-1]=reverse(trimStart(reverse(trimmed[trimmed.length-1]!),tailInset))
 const origin=trimmed[0]!.from
 return {d:[`M${origin.x} ${origin.y}`,...trimmed.map(segmentPath)].join(' '),head,tail,straight:segments.length===1&&first.kind==='line'}
}
