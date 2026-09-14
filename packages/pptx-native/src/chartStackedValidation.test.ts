import {it,expect} from 'vitest'
import {validNativeLiteralStackedBar,validNativeLiteralStackedLine} from './chartStackedValidation.js'
const fixture=()=>({profile:'literal-stacked-bar-v1',dataOrigin:'literal',grouping:'percentStacked',barDirection:'column',overlap:100,gapWidth:100,categories:['A'],series:[{index:1,order:0,values:['-1e-100'],colors:['#123456']},{index:0,order:1,values:['2e-100'],colors:['#ABCDEF']}],categoryAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},valueAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'-1',max:'1',crossesAt:'0'}})
it('admits signed literal records with independent horizontal source axes',()=>{
 const value=fixture(),before=JSON.stringify(value);expect(validNativeLiteralStackedBar(value)).toBe(true)
 value.series.reverse();const reversed=JSON.stringify(value);expect(validNativeLiteralStackedBar(value)).toBe(true);expect(JSON.stringify(value)).toBe(reversed)
 value.barDirection='bar';value.categoryAxis.position='l';value.valueAxis.position='b';expect(validNativeLiteralStackedBar(value)).toBe(true)
 expect(JSON.stringify(fixture())).toBe(before)
 const {overlap,gapWidth,barDirection,categoryAxis,valueAxis,...rest}=fixture()
 const line={...rest,profile:'literal-stacked-line-v1',xAxis:categoryAxis,yAxis:valueAxis,series:rest.series.map(({colors,...s})=>({...s,color:colors[0],widthEmu:12700}))}
 expect(validNativeLiteralStackedLine(line)).toBe(true);expect(validNativeLiteralStackedBar(line)).toBe(false)
})
it('refuses unknown keys, sparse records, invalid axes and hidden label metadata',()=>{
 const changes:((x:ReturnType<typeof fixture>)=>void)[]=[x=>{delete x.categories[0]},x=>{delete x.series[1]},x=>{x.series[0]!.order=-0},x=>{x.categoryAxis.id=-0},x=>{x.series[1]!.colors[0]='#ABCDEF\n'},x=>{Object.assign(x,{'':1})},x=>{Object.assign(x.valueAxis,{'':1})},x=>{x.valueAxis.max='-1'},x=>{x.valueAxis.crossesAt='.5'},x=>{Object.assign(x.categoryAxis,{labels:{profile:'explicit-axis-labels-v1'}})},x=>{x.dataOrigin='reference'},x=>{x.overlap=0}]
 for(const change of changes){const value=fixture();change(value);expect(validNativeLiteralStackedBar(value)).toBe(false)}
})
