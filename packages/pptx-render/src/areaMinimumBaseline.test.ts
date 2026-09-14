import {expect,it} from 'vitest'
import {validNativeLiteralArea,type NativeLiteralArea} from '@injoffice/pptx-native'
import {createNativeLiteralAreaPaths} from './literalArea.js'
import {createCartesianAreaPaths} from './cartesianAreaPaths.js'
import {chartAxisTickVectors} from './chartAxisLayout.js'
function area(grouping:NativeLiteralArea['grouping']='standard'):NativeLiteralArea{return {profile:'literal-area-v1',dataOrigin:'literal',grouping,categories:['A','B','C'],series:[{index:9,order:1,color:'#123456',values:['0','0','0']},{index:3,order:0,color:'#ABCDEF',values:['0','0','0']}],xAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:false,color:'#000000',widthEmu:12700},yAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:false,color:'#000000',widthEmu:12700,min:'-10',max:'10',crossesAt:'0'}}}
it('retains legacy zero closure and binds explicit source minimum without shifting stack tops',()=>{
 for(const grouping of ['standard','stacked','percentStacked'] as const){
  const legacy=area(grouping),bound={...legacy,sourceBaseline:{crossing:'min' as const,value:'-10'}},before=JSON.stringify(bound)
  expect(createNativeLiteralAreaPaths(legacy,100,100).filter(v=>v.seriesIndex!==undefined)).toHaveLength(0)
  const vectors=createNativeLiteralAreaPaths(bound,100,100),fills=vectors.filter(v=>v.seriesIndex!==undefined)
  expect(fills).toHaveLength(grouping==='standard'?2:1)
  for(const fill of fills)expect(fill.path.filter(p=>p.kind==='moveTo'||p.kind==='lineTo').map(p=>'y' in p?p.y:0)).toEqual(expect.arrayContaining([50,100]))
  expect(vectors.find(v=>v.axis==='x')!.path).toEqual([{kind:'moveTo',x:0,y:100},{kind:'lineTo',x:100,y:100}]);expect(JSON.stringify(bound)).toBe(before)
  const reversed={...bound,yAxis:{...bound.yAxis,orientation:'maxMin' as const}}
  expect(createNativeLiteralAreaPaths(reversed,100,100).find(v=>v.axis==='x')!.path).toEqual([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:100,y:0}])
 }
 const series=[{index:9,order:1,values:['2','2','2']},{index:3,order:0,values:['3','3','3']}]
 for(const grouping of ['stacked','percentStacked'] as const){
  const legacy=createCartesianAreaPaths(series,grouping,{min:'-10',max:'10'},100,100),bound=createCartesianAreaPaths(series,grouping,{min:'-10',max:'10',baseline:'-10'},100,100)
  expect(bound[1]).toEqual(legacy[1]);expect(bound[0]).not.toEqual(legacy[0])
 }
})
it('refuses forged baseline descriptors and retains negative stacked refusal',()=>{
 const base=area();for(const sourceBaseline of [null,{crossing:'max',value:'-10'},{crossing:'min',value:'-10.0'},{crossing:'min',value:'0'},{crossing:'min',value:'-10',extra:true}]){
  const chart={...base,sourceBaseline} as NativeLiteralArea;expect(validNativeLiteralArea(chart)).toBe(false);expect(()=>createNativeLiteralAreaPaths(chart,100,100)).toThrow()
 }
 for(const grouping of ['stacked','percentStacked'] as const){const chart={...area(grouping),sourceBaseline:{crossing:'min' as const,value:'-10'}};chart.series=[{...chart.series[0]!,values:['-1','0','1']},{...chart.series[1]!}];expect(validNativeLiteralArea(chart)).toBe(false);expect(()=>createNativeLiteralAreaPaths(chart,100,100)).toThrow()}
})
it('uses the same source minimum for existing outward horizontal ticks without changing legacy axes',()=>{
 const c=area(),axis={...c.xAxis,labels:{profile:'explicit-axis-labels-v1' as const,position:'low' as const,majorTickMark:'out' as const,style:{fontFamily:'DejaVu Sans',fontSize:1200,color:'000000',bold:false,italic:false,language:'en-US'}}},perpendicular={...c.yAxis,max:'0'},plot={x:100,y:100,cx:1000,cy:1000}
 const legacy=chartAxisTickVectors([{axis,perpendicular,horizontal:true,categories:c.categories}],plot),bound=chartAxisTickVectors([{axis,perpendicular,horizontal:true,categories:c.categories,crossingAtMinimum:true}],plot)
 expect(legacy[0]!.path[0]).toMatchObject({y:100});expect(bound[0]!.path[0]).toMatchObject({y:1100})
})
