import type { NativeEvaluatedGeometry, NativeGeometryCommand } from './types'
import type { NativeValidationIssue } from './validate'

export function validateEvaluatedGeometry(geometry:NativeEvaluatedGeometry,path:string,issues:NativeValidationIssue[]):void {
 const add=(p:string,message:string)=>issues.push({path:p,code:'native.geometry',message})
 const r=geometry.textRect
 if(r.cx<=0||r.cy<=0||!Number.isSafeInteger(r.x+r.cx)||!Number.isSafeInteger(r.y+r.cy))add(`${path}.textRect`,'invalid geometry text rectangle')
 let total=0
 const fields:Record<NativeGeometryCommand['kind'],readonly string[]>={moveTo:['x','y'],lineTo:['x','y'],quadBezierTo:['x','y','x1','y1'],cubicBezierTo:['x','y','x1','y1','x2','y2'],arcTo:['x','y','rx','ry','largeArc','clockwise'],close:[]}
 geometry.paths.forEach((part,i)=>{
  const pp=`${path}.paths[${i}]`
  total+=part.commands.length
  if(total>8192){add(pp,'geometry total command budget exceeded');return}
  let pen:readonly [number,number]|undefined,start:readonly [number,number]|undefined
  part.commands.forEach((c,j)=>{
   const cp=`${pp}.commands[${j}]`,wanted=fields[c.kind]
   if(wanted.some(k=>(c as unknown as Record<string,unknown>)[k]===undefined)||Object.keys(c).some(k=>k!=='kind'&&!wanted.includes(k))){add(cp,'command fields do not match kind');return}
   if(c.kind!=='moveTo'&&!pen){add(cp,'command precedes moveTo');return}
   if(c.kind==='close'){pen=start;return}
   if(c.kind==='arcTo'){
    if(c.rx!<=0||c.ry!<=0||c.largeArc||!arcRadiiFit(pen![0],pen![1],c.x!,c.y!,c.rx!,c.ry!))add(cp,'invalid arc radii or noncanonical long arc')
   }
   pen=[c.x!,c.y!]
   if(c.kind==='moveTo')start=pen
  })
 })
}

function arcRadiiFit(x0:number,y0:number,x1:number,y1:number,rx:number,ry:number):boolean {
 const dx=BigInt(x1)-BigInt(x0),dy=BigInt(y1)-BigInt(y0),xx=BigInt(rx)**2n,yy=BigInt(ry)**2n
 return dx*dx*yy+dy*dy*xx<=4n*xx*yy
}
