import {chartRational as r,chartRationalAdd as add,chartRationalSubtract as sub,chartRationalMultiply as mul,chartRationalDivide as div,chartRationalCompare as cmp,chartRationalDecimal as decimal,chartRationalCoordinate as coordinate,type ChartRational} from './chartRational.js'
import type {ChartRationalPoint} from './chartSegmentClip.js'
import type {ChartStackGrouping,ChartStackSeries} from './chartStacking.js'
import {createSignedMinimumAreaBands} from './signedMinimumAreaBands.js'
import type {RenderPathCommand} from './types.js'

export interface CartesianSignedAreaGeometry {readonly seriesIndex:number;readonly rings:readonly (readonly RenderPathCommand[])[]}
export interface CartesianSignedAreaScale {readonly min:string;readonly max:string;readonly crossing:'min';readonly reverseX?:boolean;readonly reverseY?:boolean}
const zero=r(0n),one=r(1n)
function interpolate(a:ChartRationalPoint,b:ChartRationalPoint,t:ChartRational):ChartRationalPoint{return {x:add(a.x,mul(sub(b.x,a.x),t)),y:add(a.y,mul(sub(b.y,a.y),t))}}
function same(a:ChartRationalPoint,b:ChartRationalPoint):boolean{return cmp(a.x,b.x)===0&&cmp(a.y,b.y)===0}
function cross(a:ChartRationalPoint,b:ChartRationalPoint,c:ChartRationalPoint):ChartRational{return sub(mul(sub(b.x,a.x),sub(c.y,b.y)),mul(sub(b.y,a.y),sub(c.x,b.x)))}
function pointKey(p:ChartRationalPoint):string{return `${p.x.numerator}/${p.x.denominator},${p.y.numerator}/${p.y.denominator}`}

/** Join closed directed walks only at an identical exact vertex. Swapping the
 * two successors at equal points preserves the multiset of directed geometric
 * edges, hence nonzero winding at every point off those edges. No connector is
 * added. Different components alone are joined, so holes and opposite winding
 * survive and a contact does not create a second cycle. Fill-only: outlines
 * are excluded by the area source profile. Exported only for independent tests. */
export function stitchExactAreaContacts(input:readonly (readonly ChartRationalPoint[])[]):ChartRationalPoint[][] {
 const nodes:{point:ChartRationalPoint;next:number;owner:number}[]=[],parents=input.map((_,i)=>i),contacts=new Map<string,number>()
 const root=(i:number):number=>{while(parents[i]!==i){parents[i]=parents[parents[i]!]!;i=parents[i]!}return i}
 input.forEach((ring,owner)=>{
  const start=nodes.length
  ring.forEach((point,i)=>nodes.push({point,next:start+(i+1)%ring.length,owner}))
 })
 for(let i=0;i<nodes.length;i++){
  const node=nodes[i]!,key=pointKey(node.point),other=contacts.get(key)
  if(other===undefined){contacts.set(key,i);continue}
  const peer=nodes[other]!,a=root(node.owner),b=root(peer.owner)
  if(a===b)continue
  const next=node.next;node.next=peer.next;peer.next=next;parents[a]=b
 }
 const seen=new Set<number>(),result:ChartRationalPoint[][]=[]
 for(let start=0;start<nodes.length;start++){
  if(seen.has(start))continue
  const ring:ChartRationalPoint[]=[]
  let current=start
  do{
   if(seen.has(current))throw new RangeError('invalid area contact walk')
   seen.add(current);ring.push(nodes[current]!.point);current=nodes[current]!.next
  }while(current!==start)
  result.push(ring)
 }
 return result
}

/** Cancel only complete, oppositely directed shared edges. Interval boundaries
 * have identical exact endpoints after clipping; no geometric epsilon or
 * approximate vertex welding is used. Point-touching rings remain separate.
 * Following a removed edge through its twin joins the two original face walks.
 */
function mergeIntervalBoundaries(polygons:readonly (readonly ChartRationalPoint[])[]):ChartRationalPoint[][]{
 const edges:{from:ChartRationalPoint;to:ChartRationalPoint;next:number;twin?:number}[]=[],open=new Map<string,number>()
 for(const polygon of polygons){
  const first=edges.length
  polygon.forEach((from,i)=>{
   const to=polygon[(i+1)%polygon.length]!,key=`${pointKey(from)}>${pointKey(to)}`,reverse=`${pointKey(to)}>${pointKey(from)}`,opposite=open.get(reverse),index=edges.length
   if(open.has(key))throw new RangeError('ambiguous area boundary')
   edges.push({from,to,next:first+(i+1)%polygon.length})
   if(opposite===undefined)open.set(key,index)
   else{edges[index]!.twin=opposite;edges[opposite]!.twin=index;open.delete(reverse)}
  })
 }
 const visited=new Set<number>(),rings:ChartRationalPoint[][]=[]
 let steps=0
 for(let start=0;start<edges.length;start++){
  if(edges[start]!.twin!==undefined||visited.has(start))continue
  const ring:ChartRationalPoint[]=[]
  let current=start
  do{
   if(visited.has(current))throw new RangeError('nonmanifold area boundary')
   visited.add(current);ring.push(edges[current]!.from)
   let next=edges[current]!.next
   while(edges[next]!.twin!==undefined){
    next=edges[edges[next]!.twin!]!.next
    if(++steps>edges.length*2)throw new RangeError('area boundary traversal budget exceeded')
   }
   current=next
  }while(current!==start)
  // A linear chain needs no intermediate source-category vertex. This removes
  // interval-only baseline points while preserving corners and source curves.
  const compact=ring.filter((point,i)=>cross(ring[(i+ring.length-1)%ring.length]!,point,ring[(i+1)%ring.length]!).numerator!==0n)
  if(compact.length>=3)rings.push(compact)
 }
 return rings
}

/** Convex interval polygons are clipped before integer quantization. Keeping
 * intervals as compound fill subpaths avoids connecting disjoint visible bands.
 * Callers must fill all same-series rings together, with no outline. */
function clipPolygon(input:readonly ChartRationalPoint[]):ChartRationalPoint[]{
 let polygon=[...input]
 for(const [axis,bound,above] of [['x',zero,true],['x',one,false],['y',zero,true],['y',one,false]] as const){
  const output:ChartRationalPoint[]=[]
  if(polygon.length===0)return output
  let previous=polygon.at(-1)!,previousInside=above?cmp(previous[axis],bound)>=0:cmp(previous[axis],bound)<=0
  for(const point of polygon){
   const inside=above?cmp(point[axis],bound)>=0:cmp(point[axis],bound)<=0
   if(inside!==previousInside)output.push(interpolate(previous,point,div(sub(bound,previous[axis]),sub(point[axis],previous[axis]))))
   if(inside)output.push(point)
   previous=point;previousInside=inside
  }
  polygon=output.filter((point,i)=>!same(point,output[(i+output.length-1)%output.length]!))
 }
 return polygon
}

/** Private area geometry, not a public chart-source projection. A complete
 * category lies at (2*i+1)/(2*n), matching the qualified between-category axis.
 * Empty/single-category/zero-height bands produce no filled rings; axes remain
 * an independent integration concern and use the explicit supplied scale. */
export function createCartesianSignedAreaPaths(series:readonly ChartStackSeries[],grouping:ChartStackGrouping,scale:CartesianSignedAreaScale,cx:number,cy:number):readonly CartesianSignedAreaGeometry[]{
 for(const extent of [cx,cy])if(!Number.isSafeInteger(extent)||extent<1||extent>281474976710655)throw new RangeError('invalid area frame')
 if((scale.reverseX!==undefined&&typeof scale.reverseX!=='boolean')||(scale.reverseY!==undefined&&typeof scale.reverseY!=='boolean'))throw new RangeError('invalid area orientation')
 const min=decimal(scale.min),max=decimal(scale.max),span=sub(max,min)
 if(cmp(min,max)>=0||cmp(min,zero)>0||cmp(max,zero)<0)throw new RangeError('invalid explicit area scale')
 const stack=createSignedMinimumAreaBands(series,grouping,scale.crossing,scale.min,scale.max)
 const y=(value:ChartRational)=>div(sub(value,min),span)
 return stack.bands.map(band=>{
  const count=band.lower.length,rings:RenderPathCommand[][]=[],polygons:ChartRationalPoint[][]=[]
  for(let point=1;point<count;point++){
   const left=r(BigInt(2*point-1),BigInt(2*count)),right=r(BigInt(2*point+1),BigInt(2*count))
   const lowerA={x:left,y:y(band.lower[point-1]!)},lowerB={x:right,y:y(band.lower[point]!)},upperA={x:left,y:y(band.upper[point-1]!)},upperB={x:right,y:y(band.upper[point]!)}
   const heightA=sub(upperA.y,lowerA.y),heightB=sub(upperB.y,lowerB.y)
   if(heightA.numerator===0n&&heightB.numerator===0n)continue
   // Signed area bands can cross either closure boundary. Split
   // exactly at the crossing so neither emitted polygon is a bow-tie.
   const pieces:ChartRationalPoint[][]=[]
   if((heightA.numerator<0n&&heightB.numerator>0n)||(heightA.numerator>0n&&heightB.numerator<0n)){
    const t=div(heightA,sub(heightA,heightB)),cross=interpolate(lowerA,lowerB,t)
    pieces.push([upperA,cross,lowerA],[cross,upperB,lowerB])
   }else pieces.push([upperA,upperB,lowerB,lowerA])
   for(const piece of pieces){
    const clipped=clipPolygon(piece)
    if(clipped.length<3)continue
    // Zero-area exact pieces must not introduce doubled or reversed zero-width
    // edges into the face graph. Quantized slivers are removed after merging.
    if(clipped.every((p,i)=>cross(clipped[(i+clipped.length-1)%clipped.length]!,p,clipped[(i+1)%clipped.length]!).numerator===0n))continue
    polygons.push(clipped)
   }
  }
  const visible=mergeIntervalBoundaries(polygons).filter(polygon=>{
    const screen=polygon.map(p=>({x:coordinate(p.x,cx,scale.reverseX??false),y:coordinate(p.y,cy,!(scale.reverseY??false))}))
    // Test individual simple lobes before stitching: a joined walk can have
    // zero signed total area while opposite-winding lobes remain visible.
    let twiceArea=0n
    for(let i=0;i<screen.length;i++){const a=screen[i]!,b=screen[(i+1)%screen.length]!;twiceArea+=BigInt(a.x)*BigInt(b.y)-BigInt(b.x)*BigInt(a.y)}
    return twiceArea!==0n
  })
  for(const polygon of stitchExactAreaContacts(visible)){
    const screen=polygon.map(p=>({x:coordinate(p.x,cx,scale.reverseX??false),y:coordinate(p.y,cy,!(scale.reverseY??false))}))
    rings.push([{kind:'moveTo',...screen[0]!},...screen.slice(1).map(p=>({kind:'lineTo' as const,...p})),{kind:'close'}])
  }
  // Preserve the established generated-area cap. Signed bands may meet both
  // clip planes; check actual topology before returning any result. The
  // preparation audit must establish the complete signed-profile bound;
  // the old nonnegative 5n proof is not asserted here.
  if(rings.length>256||rings.reduce((sum,ring)=>sum+ring.length,0)>1536)throw new RangeError(`area path budget exceeded: ${rings.length} rings, ${rings.reduce((sum,ring)=>sum+ring.length,0)} commands`)
  return {seriesIndex:band.index,rings}
 })
}
