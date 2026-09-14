import {expect,it} from 'vitest'
import {createChartRadarGeometry} from './chartRadarGeometry.js'
const radarFixture=()=>({profile:'literal-radar-v1',dataOrigin:'literal',style:'standard',categories:['A','B','C','D'],series:[{index:5,order:0,values:['2','4','6','8'],color:'#123456',widthEmu:12700}],categoryAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},valueAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'0',max:'10',crossesAt:'0'}} as const)
// Independent Decimal(90 digits) Taylor-series reference, 149 terms and a
// retained 86-digit pi constant; generated externally, no JS Math trig oracle.
const references=[[3, 0, 50000000, 0], [3, 1, 93301270, 75000000], [3, 2, 6698730, 75000000], [5, 0, 50000000, 0], [5, 1, 97552826, 34549150], [5, 2, 79389263, 90450850], [5, 4, 2447174, 34549150], [7, 0, 50000000, 0], [7, 1, 89091574, 18825510], [7, 3, 71694187, 95048443], [7, 6, 10908426, 18825510], [256, 0, 50000000, 0], [256, 1, 51227061, 15059], [256, 128, 50000000, 100000000], [256, 255, 48772939, 15059]]
it('matches independently computed high-precision angular references at the maximum frame',()=>{
 for(const [count,index,x,y]of references){
  const chart:any=structuredClone(radarFixture());chart.categories=Array.from({length:count!},(_,i)=>String(i));chart.series[0].values=Array(count!).fill('10')
  const vector=createChartRadarGeometry(chart,100000000,100000000)[0]!,point=vector.path[index!] as {x:number;y:number}
  expect([point.x,point.y]).toEqual([x,y])
  chart.categoryAxis.orientation='maxMin';const reverse=createChartRadarGeometry(chart,100000000,100000000)[0]!.path[index!] as {x:number;y:number}
  expect([reverse.x,reverse.y]).toEqual([100000000-x!,y])
 }
})
it('uses exact signed radial scaling, zero-radius endpoints and reversed radial extrema',()=>{
 const chart:any=structuredClone(radarFixture());chart.valueAxis.min='-10';chart.series[0].values=['-10','0','10','-5']
 const path=createChartRadarGeometry(chart,1000,1000)[0]!.path
 expect(path).toEqual([{kind:'moveTo',x:500,y:500},{kind:'lineTo',x:750,y:500},{kind:'lineTo',x:500,y:1000},{kind:'lineTo',x:375,y:500},{kind:'close'}])
 chart.valueAxis.orientation='maxMin';expect(createChartRadarGeometry(chart,1000,1000)[0]!.path[0]).toEqual({kind:'moveTo',x:500,y:0})
})
it('uses source value-axis paint and style-specific XML series ordering without mutating data',()=>{
 const chart:any=structuredClone(radarFixture());chart.categoryAxis={...chart.categoryAxis,deleted:false,color:'#00AA00',widthEmu:10};chart.valueAxis={...chart.valueAxis,deleted:false,color:'#AA00AA',widthEmu:20};chart.series.unshift({...chart.series[0],index:9,order:1,color:'#ABCDEF'})
 const before=JSON.stringify(chart),standard=createChartRadarGeometry(chart,1000,1000)
 expect(standard.map(v=>v.axis??v.seriesIndex)).toEqual(['value',9,5]);expect(standard[0]!.stroke.color).toBe('#AA00AA');expect(JSON.stringify(chart)).toBe(before)
 chart.style='filled';for(const s of chart.series)s.fill=s.color
 expect(createChartRadarGeometry(chart,1000,1000).map(v=>v.axis??v.seriesIndex)).toEqual([9,5,'value'])
 chart.valueAxis.deleted=true;delete chart.valueAxis.color;delete chart.valueAxis.widthEmu
 expect(createChartRadarGeometry(chart,1000,1000)).toHaveLength(2)
})
it('retains generic path limits and complete stroke plus numeric error envelopes',()=>{
 const chart:any=structuredClone(radarFixture());chart.categories=Array.from({length:256},(_,i)=>String(i));chart.series[0].values=Array(256).fill('10');chart.series=Array.from({length:16},(_,i)=>({...chart.series[0],index:i,order:i}));chart.valueAxis={...chart.valueAxis,deleted:false,color:'#123456',widthEmu:12700}
 const vectors=createChartRadarGeometry(chart,100000000,100000000)
 expect(vectors).toHaveLength(17);expect(vectors[0]!.path).toHaveLength(512);expect(vectors[1]!.path).toHaveLength(257);expect(vectors[1]!.inkBounds).toEqual({x:-6351,y:-6351,cx:100012702,cy:100012702})
 for(const bad of [NaN,Infinity,0,-1,100000001])expect(()=>createChartRadarGeometry(chart,bad,100)).toThrow()
 const outside:any=structuredClone(radarFixture());outside.series[0].values[0]='-1';expect(()=>createChartRadarGeometry(outside,100,100)).toThrow()
})
