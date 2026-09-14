import {expect,it} from 'vitest'
import {createSignedMinimumAreaBands as bands} from './signedMinimumAreaBands.js'
import {createCartesianSignedAreaPaths as paths,stitchExactAreaContacts} from './cartesianSignedAreaPaths.js'
import {chartRational as r} from './chartRational.js'
const series=(rows:string[][])=>rows.map((values,order)=>({index:10+order,order,values}))
const q=(v:{numerator:bigint;denominator:bigint})=>`${v.numerator}/${v.denominator}`
it('uses algebraic tops, absolute denominators and exact zero categories independently of closure',()=>{
 const source=series([['4','0'],['-2','0'],['3','0'],['-1','0']]),before=JSON.stringify(source)
 expect(bands(source,'stacked','min','-10','10').bands.map(b=>q(b.upper[0]!))).toEqual(['4/1','2/1','5/1','4/1'])
 const percent=bands(source,'percentStacked','min','-1','1')
 expect(percent.bands.map(b=>q(b.upper[0]!))).toEqual(['2/5','1/5','1/2','2/5'])
 expect(percent.bands.map(b=>q(b.upper[1]!))).toEqual(['0/1','0/1','0/1','0/1']);expect(percent.bands.map(b=>q(b.lower[1]!))).toEqual(['-1/1','0/1','0/1','0/1'])
 expect(bands(series([['2'],['-2']]),'percentStacked','min','-1','1').bands.map(b=>q(b.upper[0]!))).toEqual(['1/2','0/1'])
 expect(bands(series([['-2'],['-3']]),'percentStacked','min','-1','1').bands.at(-1)!.upper.map(q)).toEqual(['-1/1']);expect(JSON.stringify(source)).toBe(before)
 expect(()=>bands(source,'standard','min','-1','1')).toThrow();expect(()=>bands(source,'stacked','autoZero' as never,'-1','1')).toThrow()
})
it('preserves the 1535-command actual source witness and all625 endpoint combinations under the unchanged cap',()=>{
 let max=0
 for(const a of [-4,-1,0,1,4])for(const b of [-4,-1,0,1,4])for(const c of [-4,-1,0,1,4])for(const d of [-4,-1,0,1,4]){
  const lower=Array.from({length:256},(_,i)=>String(i%2?b:a)),delta=Array.from({length:256},(_,i)=>String(i%2?d-b:c-a)),result=paths(series([lower,delta]),'stacked',{min:'-1',max:'1',crossing:'min'},100000,100000)
  for(const shape of result){const n=shape.rings.flat().length;expect(n).toBeLessThanOrEqual(1536);max=Math.max(max,n)}
  if(a===-4&&b===4&&c===4&&d===-4){expect(result[1]!.rings).toHaveLength(1);expect(result[1]!.rings[0]).toHaveLength(1535)}
 }
 expect(max).toBe(1535)
},60000)
 it('preserves every directed edge and winding through multiple contacts and holes',()=>{
  const p=(x:number,y:number)=>({x:r(BigInt(x)),y:r(BigInt(y))})
  const outer=[p(0,0),p(4,0),p(4,4),p(0,4)]
  const hole=[p(1,1),p(1,3),p(3,3),p(3,1)]
  const clockwise=[p(4,4),p(4,6),p(6,6),p(6,4)]
  const bridge=[p(0,0),p(-2,0),p(-2,-2),p(0,-2)]
  const input=[outer,hole,clockwise,bridge],out=stitchExactAreaContacts(input)
  const key=(point:ReturnType<typeof p>)=>`${point.x.numerator},${point.y.numerator}`
  const edges=(rings:typeof input)=>rings.flatMap(ring=>ring.map((point,i)=>`${key(point)}>${key(ring[(i+1)%ring.length]!)}`)).sort()
  expect(edges(out)).toEqual(edges(input))
  expect(out).toHaveLength(2) // enclosed hole stays independent
  const winding=(rings:typeof input,x:number,y:number)=>{
   let count=0
   for(const ring of rings)for(let i=0;i<ring.length;i++){
    const a=ring[i]!,b=ring[(i+1)%ring.length]!,ax=Number(a.x.numerator),ay=Number(a.y.numerator),bx=Number(b.x.numerator),by=Number(b.y.numerator),cross=(bx-ax)*(y-ay)-(x-ax)*(by-ay)
    if(ay<=y&&by>y&&cross>0)count++
    if(ay>y&&by<=y&&cross<0)count--
   }
   return count
  }
  for(let x=-2.5;x<7;x+=.5)for(let y=-2.5;y<7;y+=.5)expect(winding(out,x+.125,y+.125)).toBe(winding(input,x+.125,y+.125))
  expect(winding(out,2,2)).toBe(0)
  expect(winding(out,2,.5)).toBe(1)
  expect(winding(out,5,5)).toBe(-1)
 })
