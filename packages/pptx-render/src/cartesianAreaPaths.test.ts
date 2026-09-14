import {describe,it,expect} from 'vitest'
import {createCartesianAreaPaths} from './cartesianAreaPaths.js'
const series=(rows:string[][])=>rows.map((values,order)=>({index:order,order,values}))
describe('exact clipped Cartesian area intervals',()=>{
 it('closes a positive band around its baseline at category centers',()=>{
  const [shape]=createCartesianAreaPaths(series([['1','1']]),'standard',{min:'0',max:'2'},400,200)
  expect(shape!.rings).toEqual([[{kind:'moveTo',x:100,y:100},{kind:'lineTo',x:300,y:100},{kind:'lineTo',x:300,y:200},{kind:'lineTo',x:100,y:200},{kind:'close'}]])
 })
 it('splits a signed crossing into triangles rather than a bow-tie',()=>{
  const [shape]=createCartesianAreaPaths(series([['-1','1']]),'standard',{min:'-1',max:'1'},400,200)
  expect(shape!.rings).toHaveLength(2)
  expect(shape!.rings.every(r=>r.length===4)).toBe(true)
  expect(shape!.rings.flat()).toContainEqual({kind:'lineTo',x:200,y:100})
 })
 it('clips a band enclosing the value scale even when every original vertex is outside',()=>{
  const shapes=createCartesianAreaPaths(series([['2','2'],['4','4']]),'stacked',{min:'0',max:'1'},400,200)
  expect(shapes[0]!.rings).toHaveLength(1)
  expect(shapes[1]!.rings).toHaveLength(0)
  expect(shapes[0]!.rings[0]).toContainEqual({kind:'lineTo',x:300,y:0})
 })
 it('keeps percent totals and zero totals predictable with explicit axes',()=>{
  const empty=createCartesianAreaPaths(series([['0','-0'],['0','0']]),'percentStacked',{min:'0',max:'1'},400,200)
  expect(empty.every(s=>s.rings.length===0)).toBe(true)
  expect(createCartesianAreaPaths(series([['1']]),'standard',{min:'0',max:'1'},400,200)[0]!.rings).toEqual([])
  const shapes=createCartesianAreaPaths(series([['1','1'],['2','2']]),'percentStacked',{min:'0',max:'1'},400,300)
  expect(shapes[0]!.rings[0]![0]).toEqual({kind:'moveTo',x:100,y:200})
  expect(shapes[1]!.rings[0]![0]).toEqual({kind:'moveTo',x:100,y:0})
 })
 it('reverses coordinates after clipping and removes rounded zero-area slivers',()=>{
  const a=createCartesianAreaPaths(series([['1','1']]),'standard',{min:'0',max:'2',reverseX:true,reverseY:true},400,200)
  expect(a[0]!.rings[0]![0]).toEqual({kind:'moveTo',x:300,y:100})
  expect(createCartesianAreaPaths(series([['1e-100','1e-100']]),'standard',{min:'0',max:'1'},400,200)[0]!.rings).toEqual([])
 })
 it('bounds full aligned source and refuses unqualified scales/negative stacks',()=>{
  const values=Array.from({length:256},(_,i)=>i%2?'1':'-1')
  const result=createCartesianAreaPaths(series([values]),'standard',{min:'-.5',max:'.5'},100000,100000)
  expect(result[0]!.rings).toHaveLength(256)
  expect(result[0]!.rings.reduce((sum,r)=>sum+r.length,0)).toBeLessThanOrEqual(1536)
  expect(result[0]!.rings.every(r=>r.length<=9)).toBe(true)
  expect(()=>createCartesianAreaPaths(series([['1','1']]),'standard',{min:'2',max:'3'},400,200)).toThrow(/scale/)
  expect(()=>createCartesianAreaPaths(series([['-1','1']]),'stacked',{min:'-1',max:'1'},400,200)).toThrow(/negative/)
  expect(()=>createCartesianAreaPaths(series([Array(257).fill('1')]),'standard',{min:'0',max:'1'},400,200)).toThrow(/point count/)
 })
 it('cancels shared interval borders but retains disjoint clipped islands',()=>{
  const flat=createCartesianAreaPaths(series([Array(256).fill('1')]),'standard',{min:'0',max:'2'},100000,100000)
  expect(flat[0]!.rings).toHaveLength(1)
  expect(flat[0]!.rings[0]).toHaveLength(5)
  const islands=createCartesianAreaPaths(series([['2','0','2','0','2'],['1','1','1','1','1']]),'stacked',{min:'0',max:'1'},10000,10000)
  expect(islands[1]!.rings).toHaveLength(2)
  expect(islands[1]!.rings.every(r=>r.length>=4)).toBe(true)
 })
 it('retains the independently interpolated filled bands after cancellation',()=>{
  let seed=19
  const random=()=>{seed=(seed*1664525+1013904223)>>>0;return seed}
  for(const grouping of ['standard','stacked','percentStacked'] as const){
   for(let fixture=0;fixture<15;fixture++){
    const count=8,rows=Array.from({length:3},()=>Array.from({length:count},()=>grouping==='standard'?Number(random()%9)-4:Number(random()%5)))
    const min=grouping==='standard'?-2:0,max=grouping==='percentStacked'?0.8:3
    const shapes=createCartesianAreaPaths(series(rows.map(row=>row.map(String))),grouping,{min:String(min),max:String(max)},65536,65536)
    const bounds=rows.map((row,order)=>row.map((value,point)=>{
     const lower=grouping==='standard'?0:rows.slice(0,order).reduce((sum,r)=>sum+r[point]!,0),upper=lower+value,total=rows.reduce((sum,r)=>sum+r[point]!,0)
     return grouping==='percentStacked'?(total===0?[0,0]:[lower/total,upper/total]):[lower,upper]
    }))
    shapes.forEach((shape,order)=>{
     for(let xi=0;xi<17;xi++)for(let yi=0;yi<13;yi++){
      const x=(xi+0.371)/17,y=(yi+0.283)/13,value=max-y*(max-min),category=x*count-0.5,left=Math.floor(category),t=category-left
      let expected=false
      if(left>=0&&left<count-1){
       const a=bounds[order]![left]!,b=bounds[order]![left+1]!,lower=a[0]!+(b[0]!-a[0]!)*t,upper=a[1]!+(b[1]!-a[1]!)*t
       // The independent ideal source test ignores only the final one-EMU
       // quantization boundary, never whole vertices or clipped regions.
       if(Math.min(Math.abs(value-lower),Math.abs(value-upper))<0.0002)continue
       expected=value>Math.min(lower,upper)&&value<Math.max(lower,upper)
      }
      const px=x*65536,py=y*65536
      let inside=false
      for(const ring of shape.rings){
       const points=ring.filter(c=>c.kind==='moveTo'||c.kind==='lineTo')
       let ringInside=false
       for(let i=0,j=points.length-1;i<points.length;j=i++){
        const a=points[i]!,b=points[j]!
        if((a.y>py)!==(b.y>py)&&px<(b.x-a.x)*(py-a.y)/(b.y-a.y)+a.x)ringInside=!ringInside
       }
       if(ringInside)inside=!inside
      }
      expect(inside,`${grouping}/${fixture}/${order}/${xi}/${yi}`).toBe(expected)
     }
    })
   }
  }
 })
})
