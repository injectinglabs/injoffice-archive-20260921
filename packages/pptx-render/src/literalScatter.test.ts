import type {NativeLiteralConnected} from '@injoffice/pptx-native'
import {expect,it} from 'vitest'
import {createNativeLiteralScatterPaths as create} from './literalScatter.js'
const fixture=():NativeLiteralConnected=>({profile:'literal-scatter-v1',dataOrigin:'literal',categories:[],series:[{index:5,order:0,xValues:['1','-1','1'],values:['-1','0','1'],color:'#0000FF',widthEmu:12700}],xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:true,min:'-1',max:'1',crossesAt:'0'},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:true,min:'-1',max:'1',crossesAt:'0'}})
it('preserves duplicate and decreasing x in source order',()=>{
 expect(create(fixture(),200,100)[0]!.path).toEqual([{kind:'moveTo',x:200,y:100},{kind:'lineTo',x:0,y:50},{kind:'lineTo',x:200,y:0}])
})
it('reverses both explicit axes and draws source axis crossings',()=>{
 const record=fixture();record.xAxis.orientation='maxMin';record.yAxis.orientation='maxMin';record.xAxis.deleted=false;record.xAxis.color='#123456';record.xAxis.widthEmu=1;record.yAxis.deleted=false;record.yAxis.color='#123456';record.yAxis.widthEmu=1
 const result=create(record,200,100)
 expect(result[0]!.path).toEqual([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:200,y:50},{kind:'lineTo',x:0,y:100}])
 expect(result[1]!.path).toEqual([{kind:'moveTo',x:0,y:50},{kind:'lineTo',x:200,y:50}]);expect(result[2]!.path).toEqual([{kind:'moveTo',x:100,y:0},{kind:'lineTo',x:100,y:100}])
})
it('refuses mismatched x/y and missing explicit numeric xscale',()=>{
 const record=fixture();expect(()=>create({...record,series:[{...record.series[0]!,xValues:['1']}]},200,100)).toThrow()
 delete record.xAxis.min;expect(()=>create(record,200,100)).toThrow()
})
