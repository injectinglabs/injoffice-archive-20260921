import {it,expect} from 'vitest'
import {validNativeLiteralArea} from './chartAreaValidation.js'
import type {NativeLiteralArea} from './chartAreaTypes.js'
const fixture=():NativeLiteralArea=>({profile:'literal-area-v1',dataOrigin:'literal',grouping:'standard',categories:['A','B'],series:[{index:3,order:0,color:'#123456',values:['-1','2']}],xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:true},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:true,min:'-2',max:'2',crossesAt:'0'}})
it('accepts exact literal records and qualified nonnegative percentages',()=>{
 const chart=fixture(),before=JSON.stringify(chart)
 expect(validNativeLiteralArea(chart)).toBe(true);expect(JSON.stringify(chart)).toBe(before)
 expect(validNativeLiteralArea({...chart,grouping:'percentStacked',series:[{...chart.series[0]!,values:['-0','1e-100']}]})).toBe(true)
})
it('rejects structural extras, foreign provenance, sparse/misaligned series and negative stacks',()=>{
 const mutations:((c:any)=>void)[]=[c=>c.dataOrigin='embedded-workbook',c=>c.profile='workbook-area-v1',c=>c.cache={},c=>c.series[0].valueReference={},c=>c.grouping='stacked',c=>c.grouping='percentStacked',c=>c.series[0].order=1,c=>c.series.push(c.series[0]),c=>c.series[0].values=['1'],c=>c.series[0].values[0]='NaN',c=>delete c.series[0],c=>delete c.categories[0],c=>c.categories=Array(257).fill('A'),c=>c.series[0].color='#abcdef',c=>c.series[0].widthEmu=1,c=>c.xAxis.min='0',c=>c.yAxis.crossesAt='1',c=>c.yAxis.min='3',c=>c.xAxis.crossAxisId=9,c=>c.yAxis.color='#123456',c=>c.yAxis.extra=true]
 for(const mutate of mutations){const chart=structuredClone(fixture());mutate(chart);expect(validNativeLiteralArea(chart)).toBe(false)}
 const emptyKey={...fixture(),'':true};expect(validNativeLiteralArea(emptyKey)).toBe(false)
 for(const value of [null,[],{},undefined])expect(validNativeLiteralArea(value)).toBe(false)
 const negativeZero=fixture();(negativeZero.series[0] as {order:number}).order=-0
 expect(validNativeLiteralArea(negativeZero)).toBe(false)
})
it('checks structural axis labels before reusing shared semantic label rules',()=>{
 const chart:any=fixture()
 chart.yAxis.deleted=false;chart.yAxis.color='#000000';chart.yAxis.widthEmu=12700
 chart.yAxis.labels={profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',majorUnit:'1',numberFormat:'0.0',style:{fontFamily:'DejaVu Sans',fontSize:1200,color:'123456',bold:false,italic:false,language:'en-US'}}
 expect(validNativeLiteralArea(chart)).toBe(true)
 const emptyStyle=structuredClone(chart);emptyStyle.yAxis.labels.style['']=true;expect(validNativeLiteralArea(emptyStyle)).toBe(false)
 for(const mutate of [(c:any)=>c.yAxis.labels.style.color='#123456',(c:any)=>c.yAxis.labels.style.fontFamily='+minor-latin',(c:any)=>c.yAxis.labels.majorUnit='0',(c:any)=>c.yAxis.labels.style.bold=1,(c:any)=>c.yAxis.labels.numberFormat='General',(c:any)=>c.yAxis.labels.style=null]){
  const copy=structuredClone(chart);mutate(copy);expect(validNativeLiteralArea(copy)).toBe(false)
 }
})
