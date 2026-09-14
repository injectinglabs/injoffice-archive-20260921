import {parseChartWorkbookRange} from './chartWorkbookRange.js'
import {it,expect} from 'vitest'
import {validNativeLiteralBar} from './literalBarValidation.js'
import {validNativeLiteralConnected} from './literalConnectedValidation.js'
import {validNativeLiteralArea} from './chartAreaValidation.js'
import {validNativeLiteralBubble} from './chartBubbleValidation.js'
import {stackedWorkbookInputs} from '../test/chartWorkbookStackedFixture.js'
import {sha} from '../test/chartWorkbookFixture.js'
import {decodeNativePptxChartWorkbookInspection as decode} from './chartWorkbookInspection.js'
import {resolveNativePptxWorkbookCharts} from './chartWorkbookCoordinator.js'
import {assertResolvedWorkbookChart} from './chartWorkbookResolution.js'

function profiles():[any,(value:any)=>boolean][] {
 const x={id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},y={id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'-10',max:'10',crossesAt:'0'}
 const series=[2,0,1].map(order=>({index:10+order,order,values:['1','2'],colors:['#123456','#ABCDEF']})),categories=['A','B']
 const line=series.map(({colors,...s})=>({...s,color:colors[0],widthEmu:12700}))
 return [[{profile:'literal-bar-v1',dataOrigin:'literal',grouping:'clustered',barDirection:'column',gapWidth:100,overlap:0,categories,series,categoryAxis:x,valueAxis:y},validNativeLiteralBar],
 [{profile:'literal-line-v1',dataOrigin:'literal',categories,series:line,xAxis:x,yAxis:y},validNativeLiteralConnected],
 [{profile:'literal-scatter-v1',dataOrigin:'literal',categories:[],series:line.map(s=>({...s,xValues:['2','1']})),xAxis:{...x,min:'-10',max:'10',crossesAt:'0'},yAxis:y},validNativeLiteralConnected],
 [{profile:'literal-area-v1',dataOrigin:'literal',grouping:'standard',categories,series:line.map(({widthEmu,...s})=>s),xAxis:x,yAxis:y},validNativeLiteralArea],
 [{profile:'literal-bubble-v1',dataOrigin:'literal',bubbleScale:100,sizeRepresents:'area',series:series.map(s=>({...s,xValues:['2','1'],sizes:['1','4']})),xAxis:{...x,min:'-10',max:'10',crossesAt:'0'},yAxis:y},validNativeLiteralBubble]]
}
it('accepts exact XML series permutations while rejecting noncanonical/missing/duplicate identity',()=>{
 for(const [source,validate] of profiles()){
  const before=JSON.stringify(source);expect(validate(source)).toBe(true)
  for(const mutate of [(c:any)=>c.series[0].order=0,(c:any)=>c.series[0].order=3,(c:any)=>c.series[0].order=-0,(c:any)=>c.series[0].order=1.5,(c:any)=>c.series[0].index=c.series[1].index,(c:any)=>c.series[0].index=-0,(c:any)=>delete c.series[1]]){
   const value=structuredClone(source);mutate(value);expect(validate(value)).toBe(false)
  }
  expect(JSON.stringify(source)).toBe(before)
 }
})
it('retains ordinary workbook series reference joins and private authority through the resolver',async()=>{
 for(const family of ['bar','line','scatter','bubble'] as const){
  const {deck,result,workbook}=stackedWorkbookInputs(family==='bar'||family==='bubble'?'bar':'line'),source=result.charts[0].source
  delete source.grouping;delete source.overlap
  source.family=family
  if(family==='scatter'||family==='bubble'){
   delete source.barDirection;delete source.gapWidth
   Object.assign(source.xAxis,{min:'-10',max:'10',crossesAt:'0'})
   const ref=(formula:string)=>({kind:'numRef',formula,range:parseChartWorkbookRange(formula),cachePresent:true})
   source.series.forEach((s:any)=>{delete s.categoryReference;s.xReference=ref('Lexical!A1:A2');if(family==='bubble')s.sizeReference=ref('Lexical!C1:C2')})
   if(family==='bubble'){source.bubbleScale=100;source.sizeRepresents='area'}
  }
  const before=JSON.stringify(result),inspection=await decode(JSON.stringify(result),deck,sha)
  let calls=0;const resolved=await resolveNativePptxWorkbookCharts(inspection,async()=>{calls++;return JSON.stringify(workbook)})
  expect(calls).toBe(1);expect(resolved.refusals).toEqual([])
  const chart=resolved.charts[0]!
  expect(chart.data.profile).toBe(`workbook-${family}-v1`)
  expect(chart.data.series.map(s=>[s.index,s.order,s.values])).toEqual([[13,3,['-1','0']],[12,2,['3','0']],[11,1,['-2','0']],[10,0,['4','0']]])
  expect(()=>assertResolvedWorkbookChart(structuredClone(chart))).toThrow()
  expect(chart.references.every(r=>r.chartCacheIgnored)).toBe(true)
  expect(JSON.stringify(result)).toBe(before)
 }
})
