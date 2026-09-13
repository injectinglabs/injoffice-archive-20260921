import type {NativeLiteralPie} from '@injoffice/pptx-native'
import type {RenderPathCommand} from './types.js'

/** Host preview policy: centered inscribed circle, clockwise from up, polygon
 * chords spanning at most two degrees; one final rounding to integer EMU.
 * Values are source literals, never formula caches. No Office layout claim. */
export function createNativeLiteralPiePaths(pie:NativeLiteralPie,cx:number,cy:number):readonly {color:string;path:readonly RenderPathCommand[]}[] {
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>281474976710655||cy>281474976710655)throw new RangeError('invalid literal pie frame')
 if(pie.profile!=='literal-pie-v1'||!Number.isInteger(pie.firstSliceAngle)||pie.firstSliceAngle<0||pie.firstSliceAngle>360||pie.values.length<1||pie.values.length>64||pie.colors.length!==pie.values.length||pie.values.some(v=>!Number.isInteger(v)||v<1||v>1e9)||pie.colors.some(c=>!/^#[0-9A-F]{6}$/.test(c)))throw new RangeError('invalid literal pie profile')
 const total=pie.values.reduce((a,b)=>a+b,0),radius=Math.min(cx,cy)/2
 let sum=0
 return pie.values.map((value,index)=>{
  const start=(pie.firstSliceAngle+sum/total*360)*Math.PI/180
  sum+=value
  const end=(pie.firstSliceAngle+sum/total*360)*Math.PI/180
  const segments=Math.max(1,Math.ceil(value/total*180))
  const path:RenderPathCommand[]=[{kind:'moveTo',x:Math.round(cx/2),y:Math.round(cy/2)}]
  for(let j=0;j<=segments;j++){
   const angle=j===segments?end:start+(end-start)*j/segments
   path.push({kind:'lineTo',x:Math.round(cx/2+radius*Math.sin(angle)),y:Math.round(cy/2-radius*Math.cos(angle))})
  }
  path.push({kind:'close'})
  return {color:pie.colors[index]!,path}
 })
}
