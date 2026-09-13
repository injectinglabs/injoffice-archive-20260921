import {expect,it} from 'vitest'
import {createNativeLiteralLinePaths as create,type LiteralConnectedRecord} from './literalLine.js'
const fixture=():LiteralConnectedRecord=>({profile:'literal-line-v1',categories:['A','B','C'],series:[{index:5,order:0,values:['0','1','0'],color:'#123456',widthEmu:12700}],xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:true},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:true,min:'0',max:'1',crossesAt:'0'}})
it('draws a continuous category line with one source round join',()=>{
 const record=fixture(),before=JSON.stringify(record),result=create(record,600,300)
 expect(result[0]).toEqual({seriesIndex:5,segmentIndices:[0,1],path:[{kind:'moveTo',x:100,y:300},{kind:'lineTo',x:300,y:0},{kind:'lineTo',x:500,y:300}],stroke:{color:'#123456',widthEmu:12700,cap:'flat',join:'round',dash:'solid'}})
 expect(JSON.stringify(record)).toBe(before)
})
it('clips true segment intersections and breaks continuity across an outside source point',()=>{
 const record=fixture();const changed={...record,series:[{...record.series[0]!,values:['0','2','0']}]}
 expect(create(changed,600,300)[0]!.path).toEqual([{kind:'moveTo',x:100,y:300},{kind:'lineTo',x:200,y:0},{kind:'moveTo',x:400,y:0},{kind:'lineTo',x:500,y:300}])
})
it('keeps a singleton series without inventing a marker',()=>{
 const record=fixture();expect(create({...record,categories:['A'],series:[{...record.series[0]!,values:['.5']}]},600,300)[0]).toMatchObject({seriesIndex:5,segmentIndices:[],path:[]})
})
it('bounds worst-case clipping to510commands without splitting a stroked series',()=>{
 const record=fixture(),categories=Array.from({length:256},(_,i)=>String(i)),values=categories.map((_,i)=>i%2?'2':'-1')
 const path=create({...record,categories,series:[{...record.series[0]!,values}]},3000000,2000000)[0]!.path
 expect(path).toHaveLength(510)
 expect(()=>create({...record,profile:'literal-scatter-v1'},600,300)).toThrow()
})
