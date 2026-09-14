import {it,expect} from 'vitest'
import {createCartesianStackedLinePaths,type CartesianStackedLineInput} from './cartesianStackedLinePaths.js'
const fixture=():CartesianStackedLineInput=>({grouping:'stacked',categories:['A','B','C'],xAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},yAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'-10',max:'10',crossesAt:'0'},series:['4','-2','3','-1'].map((value,order)=>({index:20-order,order,values:[value,value,value],color:'#123456',widthEmu:10}))})
it('emits algebraic source-order tops matching the Office mixed-sign reference',()=>{
 const chart=fixture(),before=JSON.stringify(chart),paths=createCartesianStackedLinePaths(chart,600,400)
 expect(paths.map(p=>p.path)).toEqual([120,160,100,120].map(y=>[{kind:'moveTo',x:100,y},{kind:'lineTo',x:300,y},{kind:'lineTo',x:500,y}]))
 expect(paths.map(p=>p.seriesIndex)).toEqual([20,19,18,17]);expect(JSON.stringify(chart)).toBe(before)
})
it('normalizes category-dependent totals and handles reversed axes exactly',()=>{
 const chart:CartesianStackedLineInput={...fixture(),grouping:'percentStacked',series:[{index:1,order:0,values:['1','1','0'],color:'#123456',widthEmu:10},{index:0,order:1,values:['-2','3','0'],color:'#ABCDEF',widthEmu:10}]}
 chart.xAxis={...chart.xAxis,orientation:'maxMin'};chart.yAxis={...chart.yAxis,orientation:'maxMin',min:'-1',max:'1'}
 const paths=createCartesianStackedLinePaths(chart,600,600)
 expect(paths[0]!.path).toEqual([{kind:'moveTo',x:500,y:400},{kind:'lineTo',x:300,y:375},{kind:'lineTo',x:100,y:300}])
 expect(paths[1]!.path).toEqual([{kind:'moveTo',x:500,y:200},{kind:'lineTo',x:300,y:600},{kind:'lineTo',x:100,y:300}])
})
it('uses the XML-order cumulative tops through a private axis scaffold',()=>{
 const chart=fixture();chart.series.reverse();const before=JSON.stringify(chart)
 const paths=createCartesianStackedLinePaths(chart,600,400)
 expect(paths.map(p=>p.seriesIndex)).toEqual([17,18,19,20])
 expect(paths.map(p=>p.path[0])).toEqual([220,160,200,120].map(y=>({kind:'moveTo',x:100,y})))
 expect(JSON.stringify(chart)).toBe(before)
})
it('never joins through an out-of-plot source vertex and bounds paths',()=>{
 const chart=fixture();chart.series=[{...chart.series[0]!,values:['0','20','0']}]
 const path=createCartesianStackedLinePaths(chart,600,400)[0]!.path
 expect(path).toEqual([{kind:'moveTo',x:100,y:200},{kind:'lineTo',x:200,y:0},{kind:'moveTo',x:400,y:0},{kind:'lineTo',x:500,y:200}])
 const singleton={...chart,categories:['A'],series:[{...chart.series[0]!,values:['0']}]}
 expect(createCartesianStackedLinePaths(singleton,600,400)[0]!.path).toEqual([])
})
