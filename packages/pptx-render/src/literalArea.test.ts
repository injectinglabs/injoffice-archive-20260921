import {it,expect} from 'vitest'
import {createNativeLiteralAreaPaths,type LiteralAreaInput} from './literalArea.js'
const fixture=():LiteralAreaInput=>({profile:'literal-area-v1',dataOrigin:'literal',grouping:'standard',categories:['A','B','C'],series:[{index:3,order:0,color:'#123456',values:['-1','1','-1']}],xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:false,color:'#000000',widthEmu:12700},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:false,color:'#000000',widthEmu:12700,min:'-2',max:'2',crossesAt:'0'}})
it('emits every source-series ring in one fill and preserves independent axes/source',()=>{
 const chart=fixture(),before=JSON.stringify(chart),vectors=createNativeLiteralAreaPaths(chart,600,400)
 expect(vectors).toHaveLength(3)
 expect(vectors[0]!.path.filter(c=>c.kind==='moveTo')).toHaveLength(3)
 expect(vectors[0]!.color).toBe('#123456');expect(vectors[0]!.stroke).toBeUndefined()
 expect(vectors[1]!.axis).toBe('x');expect(vectors[1]!.path[0]).toEqual({kind:'moveTo',x:0,y:200})
 expect(vectors[2]!.axis).toBe('y');expect(JSON.stringify(chart)).toBe(before)
})
it('renders zero-total axes without inventing an area and keeps authored paint order',()=>{
 const chart=fixture(),zero={...chart,grouping:'percentStacked' as const,series:[{...chart.series[0]!,values:['0','-0','0']}]}
 expect(createNativeLiteralAreaPaths(zero,600,400).map(v=>v.axis)).toEqual(['x','y'])
 const two={...chart,series:[chart.series[0]!,{...chart.series[0]!,index:9,order:1,color:'#ABCDEF'}]}
 expect(createNativeLiteralAreaPaths(two,600,400).filter(v=>v.seriesIndex!==undefined).map(v=>[v.seriesIndex,v.color])).toEqual([[3,'#123456'],[9,'#ABCDEF']])
})
it('guards literal provenance, frame/axis/style and full 256-category command budget',()=>{
 for(const mutate of [(c:any)=>c.dataOrigin='embedded-workbook',(c:any)=>c.profile='other',(c:any)=>c.xAxis.crossAxisId=9,(c:any)=>c.yAxis.widthEmu=0,(c:any)=>c.series[0].color='#ff0000',(c:any)=>c.series[0].values=['1'],(c:any)=>c.grouping='stacked']){
  const copy=structuredClone(fixture());mutate(copy);expect(()=>createNativeLiteralAreaPaths(copy,600,400)).toThrow()
 }
 expect(()=>createNativeLiteralAreaPaths(fixture(),0,400)).toThrow()
 const chart=fixture(),values=Array.from({length:256},(_,i)=>i%2?'1':'-1')
 const large={...chart,categories:values,series:[{...chart.series[0]!,values}],yAxis:{...chart.yAxis,min:'-.5',max:'.5'}}
 const fill=createNativeLiteralAreaPaths(large,1000000,1000000)[0]!
 expect(fill.path.length).toBe(1280);expect(fill.path.length).toBeLessThanOrEqual(1536)
})
