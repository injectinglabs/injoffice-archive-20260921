import {it,expect} from 'vitest'
import {validateStackedChartGeometry} from './chartStackedGeometryValidation.js'
const fixture=()=>({grouping:'stacked',overlap:100,barDirection:'column',gapWidth:100,categories:['A'],categoryAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:true},valueAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:false,color:'#123456',widthEmu:1,min:'-10',max:'10',crossesAt:'0'},series:[{index:0,order:0,values:['1'],colors:['#123456']},{index:1,order:1,values:['-1'],colors:['#ABCDEF']}]})
it('rejects sparse data, newline RGB, negative-zero identities and unknown record keys',()=>{
 const mutations:((v:ReturnType<typeof fixture>)=>void)[]=[
  v=>{delete v.categories[0]},v=>{delete v.series[1]},v=>{delete v.series[1]!.values[0]},v=>{delete v.series[1]!.colors[0]},
  v=>{v.series[1]!.colors[0]='#123456\n'},v=>{v.valueAxis.color='#123456\n'},v=>{v.categoryAxis.id=-0},v=>{v.valueAxis.crossAxisId=-0},v=>{v.gapWidth=-0},v=>{v.series[0]!.order=-0},v=>{Object.assign(v,{'':true})},v=>{Object.assign(v.valueAxis,{'':true})},v=>{Object.assign(v.series[0]!,{unexpected:1})},
 ]
 for(const mutate of mutations){const value=fixture();mutate(value);expect(()=>validateStackedChartGeometry(value,'bar')).toThrow()}
 expect(()=>validateStackedChartGeometry(fixture(),'bar')).not.toThrow()
})
it('applies full RGB and closed series records to line strokes too',()=>{
 const bar=fixture(),line={grouping:'percentStacked',categories:bar.categories,xAxis:bar.categoryAxis,yAxis:bar.valueAxis,series:bar.series.map(({colors,...s})=>({...s,color:colors[0],widthEmu:1}))}
 expect(()=>validateStackedChartGeometry(line,'line')).not.toThrow()
 line.series[1]!.color='#ABCDEF\n';expect(()=>validateStackedChartGeometry(line,'line')).toThrow()
})
