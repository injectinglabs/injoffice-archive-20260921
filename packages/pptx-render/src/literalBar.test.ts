import {describe,expect,it} from 'vitest'
import {createNativeLiteralBarPaths} from './literalBar.js'
import type {NativeLiteralBar as LiteralBarRecord} from '@injoffice/pptx-native'
function chart():LiteralBarRecord{return {profile:'literal-bar-v1',barDirection:'column',grouping:'clustered',dataOrigin:'literal',gapWidth:100,overlap:0,categories:['A','B'],series:[{index:7,order:0,values:['10','-5'],colors:['#123456','#ABCDEF']},{index:9,order:1,values:['20','0'],colors:['#000000','#FFFFFF']}],categoryAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},valueAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'-10',max:'20',crossesAt:'0'}}}
const rectangle=(x:number,y:number,right:number,bottom:number)=>[{kind:'moveTo',x,y},{kind:'lineTo',x:right,y},{kind:'lineTo',x:right,y:bottom},{kind:'lineTo',x,y:bottom},{kind:'close'}]
describe('literal clustered bars',()=>{
 it('derives gap as one bar width, retains signed/zero source identities',()=>{
  const result=createNativeLiteralBarPaths(chart(),600,300)
  expect(result).toHaveLength(4)
  expect(result[0]).toEqual({seriesIndex:7,pointIndex:0,color:'#123456',path:rectangle(50,100,150,200)})
  expect(result[1]!.path).toEqual(rectangle(350,200,450,250))
  expect(result[2]!.path).toEqual(rectangle(150,0,250,200))
  expect(result[3]!.path).toEqual(rectangle(450,200,550,200))
 })
 it('transposes the value axis and places increasing bar categories bottom to top',()=>{
  const c=chart();c.barDirection='bar';c.categoryAxis.position='l';c.valueAxis.position='b'
  const result=createNativeLiteralBarPaths(c,300,600)
  expect(result[0]!.path).toEqual(rectangle(100,450,200,550))
  expect(result[1]!.path).toEqual(rectangle(50,150,100,250))
  c.categoryAxis.orientation='maxMin';c.valueAxis.orientation='maxMin'
  expect(createNativeLiteralBarPaths(c,300,600)[0]!.path).toEqual(rectangle(100,50,200,150))
 })
 it('clips bars to the explicit scale while drawing only explicitly visible axis lines',()=>{
  const c=chart();c.series[0]!.values=['1e100','-1e100'];c.categoryAxis.deleted=false;c.categoryAxis.color='#111111';c.categoryAxis.widthEmu=12700;c.valueAxis.deleted=false;c.valueAxis.color='#222222';c.valueAxis.widthEmu=25400
  const result=createNativeLiteralBarPaths(c,600,300)
  expect(result[0]!.path).toEqual(rectangle(50,0,150,200));expect(result[1]!.path).toEqual(rectangle(350,200,450,300))
  expect(result[4]).toMatchObject({axis:'category',path:[{kind:'moveTo',x:0,y:200},{kind:'lineTo',x:600,y:200}]})
  expect(result[5]).toMatchObject({axis:'value',path:[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:0,y:300}]})
 })
 it('refuses invalid profile, point, axis, and frame combinations',()=>{
  for(const mutate of [(c:LiteralBarRecord)=>{c.series[0]!.values=['1']},(c:LiteralBarRecord)=>{c.series[0]!.values=['NaN','1']},(c:LiteralBarRecord)=>{c.series[1]!.index=7},(c:LiteralBarRecord)=>{c.valueAxis.crossesAt='1'},(c:LiteralBarRecord)=>{c.categoryAxis.color='#000000'},(c:LiteralBarRecord)=>{c.valueAxis.min='1'},(c:LiteralBarRecord)=>{c.gapWidth=501}]){const c=chart();mutate(c);expect(()=>createNativeLiteralBarPaths(c,600,300)).toThrow()}
  expect(()=>createNativeLiteralBarPaths(chart(),1,300)).toThrow(/distinct bars/)
 })
})
