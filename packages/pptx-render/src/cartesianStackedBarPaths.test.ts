import {it,expect} from 'vitest'
import {createCartesianStackedBarPaths,type CartesianStackedBarInput} from './cartesianStackedBarPaths.js'
const fixture=():CartesianStackedBarInput=>({grouping:'stacked',overlap:100,barDirection:'column',gapWidth:100,categories:['A'],categoryAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:false,color:'#111111',widthEmu:10},valueAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:false,color:'#111111',widthEmu:10,min:'-10',max:'10',crossesAt:'0'},series:['4','-2','3','-1'].map((value,order)=>({index:order+10,order,values:[value],colors:['#123456']}))})
it('paints source-ordered diverging stack rectangles and independent axes',()=>{
 const chart=fixture(),before=JSON.stringify(chart),paths=createCartesianStackedBarPaths(chart,400,400)
 expect(paths.slice(0,4).map(p=>p.path.slice(0,3))).toEqual([[120,200],[200,240],[60,120],[240,260]].map(([top,bottom])=>[{kind:'moveTo',x:100,y:top},{kind:'lineTo',x:300,y:top},{kind:'lineTo',x:300,y:bottom}]))
 expect(paths.slice(0,4).map(p=>p.seriesIndex)).toEqual([10,11,12,13]);expect(paths.slice(4).map(p=>p.axis)).toEqual(['category','value']);expect(JSON.stringify(chart)).toBe(before)
})
it('handles horizontal bars and both reversed axes without changing source stacking',()=>{
 const chart=fixture();chart.barDirection='bar';chart.categoryAxis={...chart.categoryAxis,position:'l',orientation:'maxMin'};chart.valueAxis={...chart.valueAxis,position:'b',orientation:'maxMin'}
 const paths=createCartesianStackedBarPaths(chart,400,400)
 expect(paths[0]!.path.slice(0,3)).toEqual([{kind:'moveTo',x:120,y:100},{kind:'lineTo',x:200,y:100},{kind:'lineTo',x:200,y:300}])
 expect(paths[1]!.path[0]).toEqual({kind:'moveTo',x:200,y:100})
})
it('keeps XML paint sequence distinct from numeric order metadata',()=>{
 const chart=fixture();chart.series.reverse();const before=JSON.stringify(chart)
 const paths=createCartesianStackedBarPaths(chart,400,400)
 expect(paths.slice(0,4).map(p=>p.seriesIndex)).toEqual([13,12,11,10])
 expect(paths.slice(0,4).map(p=>p.path[0])).toEqual([200,140,220,60].map(y=>({kind:'moveTo',x:100,y})))
 expect(JSON.stringify(chart)).toBe(before)
})
it('clips exact percent endpoints and leaves zero categories as axes only',()=>{
 const chart:CartesianStackedBarInput={...fixture(),grouping:'percentStacked'};chart.valueAxis={...chart.valueAxis,min:'-.5',max:'.5'}
 expect(createCartesianStackedBarPaths(chart,400,400)[2]!.path[0]).toEqual({kind:'moveTo',x:100,y:0})
 const zero={...chart,series:chart.series.map(s=>({...s,values:['0']}))}
 expect(createCartesianStackedBarPaths(zero,400,400).map(p=>p.axis)).toEqual(['category','value'])
 expect(()=>createCartesianStackedBarPaths({...chart,categories:['A','B'],series:chart.series.map(s=>({...s,colors:['#123456','#123456']}))},400,400)).toThrow()
})
