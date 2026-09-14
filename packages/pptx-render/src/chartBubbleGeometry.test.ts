import {expect,it} from 'vitest'
import {createChartBubbleGeometry as geometry,type ChartBubbleSeries,type ChartBubbleScale} from './chartBubbleGeometry.js'
const scale:ChartBubbleScale={xMin:'0',xMax:'10',yMin:'0',yMax:'10',bubbleScale:100,sizeRepresents:'area',sizingPolicy:'plot-minor-radius-v1'}
const series=(sizes=['1','4']):ChartBubbleSeries[]=>[{index:7,order:0,xValues:['2','8'],values:['3','7'],sizes}]
it('preserves exact center mapping and meaningful area/width/scale sizing',()=>{
 const source=series(),before=JSON.stringify(source),area=geometry(source,scale,1000,800)
 expect(area.map(p=>[p.centerX,p.centerY,p.radius])).toEqual([[200,560,40],[800,240,80]])
 expect(geometry(source,{...scale,sizeRepresents:'w'},1000,800).map(p=>p.radius)).toEqual([20,80])
 expect(geometry(source,{...scale,bubbleScale:300},1000,800).map(p=>p.radius)).toEqual([120,240])
 expect(geometry(source,{...scale,reverseX:true,reverseY:true},1000,800).map(p=>[p.centerX,p.centerY])).toEqual([[800,240],[200,560]])
 expect(area.every(p=>p.path.length===4)).toBe(true);expect(JSON.stringify(source)).toBe(before)
})
it('uses one global size maximum across series and retains source identity/order',()=>{
 const a=series(['1','1'])[0]!,b={...a,index:2,order:1,sizes:['4','1']}
 expect(geometry([a,b],scale,1000,1000).map(p=>[p.seriesIndex,p.pointIndex,p.radius])).toEqual([[7,0,50],[7,1,50],[2,0,100],[2,1,50]])
})
it('keeps intersecting outside-center circles, culls huge outside values and keeps no invented minimum',()=>{
 const s=series()[0]!
 const source=[{...s,xValues:['-0.5','1e100'],values:['5','5'],sizes:['4','4']}]
 const result=geometry(source,scale,1000,1000)
 expect(result).toHaveLength(1);expect(result[0]!.centerX).toBe(-50);expect(result[0]!.radius).toBe(100)
 expect(geometry(series(['1e-100','4']),scale,1000,1000).map(p=>p.pointIndex)).toEqual([1])
 expect(geometry(series(['0','0']),scale,1000,1000)).toEqual([])
 expect(geometry(series(['0','1']),{...scale,bubbleScale:0},1000,1000)).toEqual([])
 expect(geometry(series(['0','1']),scale,1000,1000).map(p=>p.pointIndex)).toEqual([1])
})
it('guards dense aligned values, nonnegative sizes and canonical source indices',()=>{
 const invalid=[{...series()[0]!,sizes:['-1','1']},{...series()[0]!,sizes:['1']},{...series()[0]!,order:-0},{...series()[0]!,xValues:['NaN','1']},{...series()[0]!,values:new Array(2)}]
 for(const s of invalid)expect(()=>geometry([s],scale,1000,1000)).toThrow()
 expect(()=>geometry(series(),{...scale,bubbleScale:301},1000,1000)).toThrow()
 expect(()=>geometry(series(),{...scale,xMax:'0'},1000,1000)).toThrow()
 expect(()=>geometry(series(),scale,0,1000)).toThrow()
 expect(()=>geometry(new Array(1),scale,1000,1000)).toThrow()
})
it('has at most4096 independent four-command paths without global path-budget expansion',()=>{
 const values=Array(256).fill('5'),sizes=Array(256).fill('1')
 const source=Array.from({length:16},(_,i)=>({index:i,order:i,xValues:values,values,sizes}))
 const result=geometry(source,scale,1000,1000)
 expect(result).toHaveLength(4096);expect(result.reduce((n,p)=>n+p.path.length,0)).toBe(16384)
 expect(()=>geometry([...source,{...source[0]!,index:16,order:16}],scale,1000,1000)).toThrow()
})
