import {expect,it} from 'vitest'
import type {NativeLiteralBarAxis} from '@injoffice/pptx-native'
import type {RenderTextBodyNode} from './types.js'
import {layoutChartAxes,chartAxisTickVectors,type ChartAxisLabelInput} from './chartAxisLayout.js'
const body={} as RenderTextBodyNode
const measure=async()=>({body,bounds:{x:-10,y:-20,cx:100000,cy:50000}})
function axes():ChartAxisLabelInput[]{
 const style={fontFamily:'Source',fontSize:1200,color:'000000',bold:false,italic:false,language:'en-US'}
 const x:NativeLiteralBarAxis={id:1,crossAxisId:2,position:'b',orientation:'minMax',deleted:false,color:'#000000',widthEmu:100,labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'out',style}}
 const y:NativeLiteralBarAxis={id:2,crossAxisId:1,position:'l',orientation:'minMax',deleted:false,min:'0',max:'2',crossesAt:'0',color:'#000000',widthEmu:100,labels:{profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'out',style,majorUnit:'1',numberFormat:'0'}}
 return [{axis:x,perpendicular:y,horizontal:true,categories:['A','B']},{axis:y,perpendicular:x,horizontal:false}]
}
it('derives low/high placement from perpendicular value order across all reversal combinations',async()=>{
 for(const xr of ['minMax','maxMin'] as const)for(const yr of ['minMax','maxMin'] as const)for(const position of ['low','high'] as const){
  const inputs=axes();inputs[0]!.axis.orientation=xr;inputs[1]!.axis.orientation=yr
  for(const input of inputs)input.axis.labels!.position=position
  const result=await layoutChartAxes(inputs,1000000,1000000,measure),p=result.plot
  const horizontal=result.labels.slice(0,2),vertical=result.labels.slice(2)
  for(const l of horizontal)expect(l.bounds.y>p.y+p.cy).toBe((position==='low')===(yr==='minMax'))
  for(const l of vertical)expect(l.bounds.x<p.x).toBe((position==='low')===(xr==='minMax'))
  expect(horizontal[0]!.bounds.x<horizontal[1]!.bounds.x).toBe(xr==='minMax')
  expect(vertical[0]!.bounds.y>vertical[1]!.bounds.y).toBe(yr==='minMax')
  expect(result.labels[0]!.x).toBe(result.labels[0]!.bounds.x+10)
 }
})
it('places out ticks at the actual axis edge independently of label side',async()=>{
 const input=axes();input[0]!.axis.labels!.position='high';input[1]!.axis.labels!.position='high'
 const plot={x:100000,y:100000,cx:800000,cy:800000},ticks=chartAxisTickVectors(input,plot)
 expect(ticks[0]!.path).toEqual([{kind:'moveTo',x:300000,y:900000},{kind:'lineTo',x:300000,y:938100}])
 expect(ticks[2]!.path).toEqual([{kind:'moveTo',x:100000,y:900000},{kind:'lineTo',x:61900,y:900000}])
 input[1]!.axis.min='-2';input[1]!.axis.max='0'
 expect(chartAxisTickVectors(input,plot)[0]!.path).toEqual([{kind:'moveTo',x:300000,y:100000},{kind:'lineTo',x:300000,y:61900}])
})
it('refuses overlapping labels and insufficient frame instead of dropping source labels',async()=>{
 await expect(layoutChartAxes(axes(),100,100,measure)).rejects.toThrow()
 const inputs=axes();inputs[0]!.categories=Array.from({length:256},(_,i)=>String(i))
 await expect(layoutChartAxes(inputs,1000000,1000000,measure)).rejects.toThrow(/overlap/)
})

it('reserves measured cross-axis clearance for numeric endpoint labels at the same corner',async()=>{
 const inputs=axes();delete inputs[0]!.categories
 Object.assign(inputs[0]!.axis,{min:'0',max:'2',crossesAt:'0'})
 Object.assign(inputs[0]!.axis.labels!,{position:'high',majorUnit:'1',numberFormat:'0'})
 inputs[1]!.axis.labels!.position='high'
 const wide=async()=>({body,bounds:{x:0,y:0,cx:300000,cy:200000}})
 const result=await layoutChartAxes(inputs,2000000,2000000,wide)
 expect(result.labels).toHaveLength(6)
 // The top max label and right max label are separated in both projections.
 const top=result.labels[2]!.bounds,right=result.labels[5]!.bounds
 expect(top.x+top.cx<right.x).toBe(true)
 expect(top.y+top.cy<right.y).toBe(true)
})
