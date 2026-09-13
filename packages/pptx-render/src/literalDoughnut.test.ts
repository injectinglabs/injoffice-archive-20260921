import {expect,it} from 'vitest'
import {createNativeLiteralDoughnutPaths} from './literalDoughnut.js'
const ring={profile:'literal-doughnut-v1' as const,firstSliceAngle:0,holeSize:50,values:[1,3],colors:['#FF0000','#0000FF']}
// Independent ray crossing tests actual filled polygons, including a full ring.
function inside(path:ReturnType<typeof createNativeLiteralDoughnutPaths>[number]['path'],x:number,y:number){
 const points=path.filter(p=>p.kind==='moveTo'||p.kind==='lineTo');let result=false
 for(let i=0,j=points.length-1;i<points.length;j=i++){
  const a=points[i]!,b=points[j]!
  if((a.y>y)!==(b.y>y)&&x<(b.x-a.x)*(y-a.y)/(b.y-a.y)+a.x)result=!result
 }
 return result
}
it('paints annular sectors with source proportions and a transparent center',()=>{
 const paths=createNativeLiteralDoughnutPaths(ring,200,100)
 expect(paths.map(p=>p.color)).toEqual(ring.colors)
 expect(paths[0]!.path[0]).toEqual({kind:'moveTo',x:100,y:0})
 expect(paths[0]!.path[45]).toEqual({kind:'lineTo',x:150,y:50})
 expect(paths[0]!.path[46]).toEqual({kind:'lineTo',x:125,y:50})
 expect(paths[0]!.path.at(-2)).toEqual({kind:'lineTo',x:100,y:25})
 expect(inside(paths[0]!.path,125,25)).toBe(true)
 expect(paths.some(p=>p.color==='#0000FF'&&inside(p.path,75,75))).toBe(true)
 for(const path of paths){expect(inside(path.path,100,50)).toBe(false);expect(path.path.length).toBeLessThanOrEqual(363)}
 expect(createNativeLiteralDoughnutPaths({...ring,firstSliceAngle:90},200,100)[0]!.path[0]).toEqual({kind:'moveTo',x:150,y:50})
})
it('retains the hole for a single full ring at hole bounds and odd frame sizes',()=>{
 for(const holeSize of [10,50,90]){
  const paths=createNativeLiteralDoughnutPaths({...ring,holeSize,values:[1],colors:['#FF0000']},1001,1001)
  expect(paths).toHaveLength(1)
  expect(paths.every(p=>p.path.length===363)).toBe(true)
  expect(paths.some(p=>inside(p.path,500.5,500.5))).toBe(false)
  expect(paths.some(p=>inside(p.path,500.5,10))).toBe(true)
  for(const p of paths.flatMap(p=>p.path))if(p.kind==='lineTo'||p.kind==='moveTo')expect(Number.isSafeInteger(p.x)&&Number.isSafeInteger(p.y)).toBe(true)
 }
 const many=createNativeLiteralDoughnutPaths({...ring,values:Array(64).fill(1),colors:Array(64).fill('#000000')},1000,2000)
 expect(many).toHaveLength(64)
 expect(many.reduce((n,p)=>n+p.path.length,0)).toBeLessThanOrEqual(616)
})
it('rejects malformed profiles and frames that cannot retain a hole',()=>{
 for(const bad of [{...ring,holeSize:9},{...ring,holeSize:91},{...ring,holeSize:50.5},{...ring,firstSliceAngle:361},{...ring,values:[0,3]},{...ring,values:[NaN,3]},{...ring,values:[1.2,3]},{...ring,colors:['#ff0000','#0000FF']},{...ring,values:Array(65).fill(1),colors:Array(65).fill('#000000')}])expect(()=>createNativeLiteralDoughnutPaths(bad,100,100)).toThrow()
 for(const [cx,cy] of [[1,1],[1,100],[Infinity,100],[2**49,100]])expect(()=>createNativeLiteralDoughnutPaths(ring,cx!,cy!)).toThrow()
 expect(()=>createNativeLiteralDoughnutPaths({...ring,holeSize:10},19,19)).toThrow(/retain/)
 expect(createNativeLiteralDoughnutPaths({...ring,holeSize:10},20,20)).toHaveLength(2)
})
