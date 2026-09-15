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
 /** Ends whose terminal segment is shorter than its inset; the shaft is left
  * untrimmed there because the filled arrowhead covers the stub. */
 readonly skippedInsets:readonly ('head'|'tail')[]
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
 * Straight segments trim exactly; curves use a bounded bisection on the chord
 * distance and a de Casteljau split. A segment shorter than the inset is
 * returned untrimmed (`trimmed:false`): the arrowhead covers it entirely and
 * inventing a reversed or empty shaft would be worse than the stub. */
function trimStart(segment:Segment,inset:number):{readonly segment:Segment;readonly trimmed:boolean} {
 if(inset<=0)return {segment,trimmed:true}
 const chord=distance(segment.from,segment.to)
 if(chord<=inset)return {segment,trimmed:false}
 if(segment.kind==='line')return {segment:tailPortion(segment,inset/chord),trimmed:true}
 let lo=0,hi=1
 for(let i=0;i<SPLIT_ITERATIONS;i++){const mid=(lo+hi)/2;if(distance(at(segment,mid),segment.from)<inset)lo=mid;else hi=mid}
 return {segment:tailPortion(segment,hi),trimmed:true}
}

/** First point of a segment that differs from its start, or undefined for a
 * fully degenerate (zero-length) segment. */
function firstTangent(segment:Segment):Point|undefined {
 const candidates=segment.kind==='line'?[segment.to]:segment.kind==='quad'?[segment.c1,segment.to]:[segment.c1,segment.c2,segment.to]
 return candidates.find(point=>!same(point,segment.from))
}

/** Outward direction at the path start: tip minus the first point along the
 * path that differs from it, skipping zero-length leading segments. A path
 * whose every point coincides yields a zero vector, which previewArrow refuses. */
function outwardDirection(tip:Point,segments:readonly Segment[]):Point {
 for(const segment of segments){
  const tangent=firstTangent(segment)
  if(tangent)return {x:tip.x-tangent.x,y:tip.y-tangent.y}
 }
 return {x:0,y:0}
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
 const reversed=segments.map(reverse).reverse()
 const head={tip:first.from,direction:outwardDirection(first.from,segments)}
 const tail={tip:last.to,direction:outwardDirection(last.to,reversed)}
 const skippedInsets:('head'|'tail')[]=[]
 const trimmed=[...segments]
 const headTrim=trimStart(trimmed[0]!,previewArrowShaftInset(headEnd,strokeWidth))
 trimmed[0]=headTrim.segment
 if(!headTrim.trimmed)skippedInsets.push('head')
 const tailTrim=trimStart(reverse(trimmed[trimmed.length-1]!),previewArrowShaftInset(tailEnd,strokeWidth))
 trimmed[trimmed.length-1]=reverse(tailTrim.segment)
 if(!tailTrim.trimmed)skippedInsets.push('tail')
 const origin=trimmed[0]!.from
 return {d:[`M${origin.x} ${origin.y}`,...trimmed.map(segmentPath)].join(' '),head,tail,straight:segments.length===1&&first.kind==='line',skippedInsets}
}
