import type { NativeEvaluatedGeometry } from '@injoffice/pptx-native'
import { PPTX_RENDER_LIMITS, RenderCompileError, type RenderPathCommand, type RenderRect } from './types.js'

/** The native validator checks structure and arc consistency before compilation.
 * This boundary additionally honors the caller's narrower coordinate budget. */
export function evaluatedGeometryPaths(geometry:NativeEvaluatedGeometry,check:(value:number,path:string)=>void,path:string,checkWorld:(bounds:RenderRect)=>void){
 let total=0
 for(const [key,value] of Object.entries(geometry.textRect))check(value,`${path}.textRect.${key}`)
 checkWorld(geometry.textRect)
 return geometry.paths.map((part,i)=>{
  total+=part.commands.length
  if(part.commands.length>PPTX_RENDER_LIMITS.maxPathCommands||total>8192)throw new RenderCompileError('render.pathBudget',`${path}.paths[${i}]`,'geometry path budget exceeded')
  let minX=Infinity,minY=Infinity,maxX=-Infinity,maxY=-Infinity,penX=0,penY=0,startX=0,startY=0
  const include=(x:number,y:number)=>{check(x,`${path}.paths[${i}].envelope.x`);check(y,`${path}.paths[${i}].envelope.y`);minX=Math.min(minX,x);minY=Math.min(minY,y);maxX=Math.max(maxX,x);maxY=Math.max(maxY,y)}
  const commands=part.commands.map((command,j)=>{
   for(const [key,value] of Object.entries(command))if(typeof value==='number')check(value,`${path}.paths[${i}].commands[${j}].${key}`)
   for(const [x,y] of [[command.x,command.y],[command.x1,command.y1],[command.x2,command.y2]])if(x!==undefined&&y!==undefined)include(x,y)
   // A whole ellipse is bounded by a two-radius box around its start point.
   // Use a conservative envelope so excursions outside the frame cannot evade
   // cumulative group/rotation limits; no geometric coordinates are clamped.
   if(command.kind==='arcTo'){include(penX-2*command.rx!,penY-2*command.ry!);include(penX+2*command.rx!,penY+2*command.ry!)}
   if(command.kind==='close'){penX=startX;penY=startY}else{penX=command.x!;penY=command.y!;if(command.kind==='moveTo'){startX=penX;startY=penY}}
   return {...command} as RenderPathCommand
  })
  checkWorld({x:minX,y:minY,cx:maxX-minX,cy:maxY-minY})
  return {path:commands,fillMode:part.fillMode,stroke:part.stroke}
 })
}
