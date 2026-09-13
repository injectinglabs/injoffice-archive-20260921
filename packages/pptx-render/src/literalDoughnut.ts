import type {NativeLiteralDoughnut} from '@injoffice/pptx-native'
import type {RenderPathCommand} from './types.js'

/** Read-only host preview: centered circular annulus, clockwise from up,
 * two-degree polygon chords and final integer EMU rounding. Not Office layout. */
export function createNativeLiteralDoughnutPaths(chart:NativeLiteralDoughnut,cx:number,cy:number):readonly {color:string;path:readonly RenderPathCommand[]}[] {
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>281474976710655||cy>281474976710655)throw new RangeError('invalid literal doughnut frame')
 if(chart.profile!=='literal-doughnut-v1'||!Number.isInteger(chart.firstSliceAngle)||chart.firstSliceAngle<0||chart.firstSliceAngle>360||!Number.isInteger(chart.holeSize)||chart.holeSize<10||chart.holeSize>90||chart.values.length<1||chart.values.length>64||chart.colors.length!==chart.values.length||chart.values.some(v=>!Number.isInteger(v)||v<1||v>1e9)||chart.colors.some(c=>!/^#[0-9A-F]{6}$/.test(c)))throw new RangeError('invalid literal doughnut profile')
 const total=chart.values.reduce((a,b)=>a+b,0),radius=Math.min(cx,cy)/2,inner=radius*chart.holeSize/100
 // A collapsed hole or ring would misrepresent the source chart after rounding.
 if(inner<1||radius-inner<1)throw new RangeError('literal doughnut frame cannot retain its hole and ring')
 let sum=0
 return chart.values.map((value,index)=>{
  const start=chart.firstSliceAngle+sum/total*360
  sum+=value
  const end=chart.firstSliceAngle+sum/total*360
  const segments=Math.max(1,Math.ceil(value/total*180))
  const point=(r:number,j:number)=>{
   const degrees=((j===segments?end:start+(end-start)*j/segments)%360+360)%360
   const angle=degrees*Math.PI/180
   const sin=degrees===0||degrees===180?0:degrees===90?1:degrees===270?-1:Math.sin(angle)
   const cos=degrees===90||degrees===270?0:degrees===0?1:degrees===180?-1:Math.cos(angle)
   return {x:Math.round(cx/2+r*sin),y:Math.round(cy/2-r*cos)}
  }
  const path:RenderPathCommand[]=[{kind:'moveTo',...point(radius,0)}]
  for(let j=1;j<=segments;j++)path.push({kind:'lineTo',...point(radius,j)})
  for(let j=segments;j>=0;j--)path.push({kind:'lineTo',...point(inner,j)})
  path.push({kind:'close'})
  return {color:chart.colors[index]!,path}
 })
}
