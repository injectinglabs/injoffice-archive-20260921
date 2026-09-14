import {expect,it} from 'vitest'
import {readFileSync} from 'node:fs'
import {fixture,sha} from '../test/chartWorkbookFixture.js'
import {parseChartWorkbookRange} from './chartWorkbookRange.js'
import {decodeNativePptxChartWorkbookInspection as decode} from './chartWorkbookInspection.js'
import {resolveNativePptxWorkbookCharts} from './chartWorkbookCoordinator.js'

function inputs(){
 const {deck,result:base}=fixture(),result:any=base,c=base.charts[0]!
 const ref={kind:'numRef',formula:'Lexical!A1',range:parseChartWorkbookRange('Lexical!A1'),cachePresent:true}
 result.charts[0].source={family:'bubble',bubbleScale:0,sizeRepresents:'area',plotVisibleOnly:false,xAxis:{...c.source.xAxis,min:'-10',max:'10',crossesAt:'0'},yAxis:c.source.yAxis,series:[{index:0,order:0,xReference:ref,valueReference:ref,sizeReference:ref,colors:['#FF0000']}]}
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8'))
 workbook.source.package_sha256=`sha256:${c.workbook.sha256}`;workbook.revision=`rev:${c.workbook.sha256}`
 return {deck,result,workbook}
}
it('retains authoritative numeric size lexemes, zero scale and separate workbook identity',async()=>{
 const {deck,result,workbook}=inputs(),inspection=await decode(JSON.stringify(result),deck,sha)
 const resolved=await resolveNativePptxWorkbookCharts(inspection,async()=>JSON.stringify(workbook))
 expect(resolved.refusals).toEqual([]);expect(resolved.charts).toHaveLength(1)
 const data=resolved.charts[0]!.data
 expect(data.profile).toBe('workbook-bubble-v1')
 if(data.profile!=='workbook-bubble-v1')throw Error('wrong family')
 expect(data.bubbleScale).toBe(0);expect(data.series[0]!.sizes).toEqual(['001.2300']);expect(data.series[0]!.xValues).toEqual(['001.2300'])
 expect(resolved.charts[0]!.references).toHaveLength(1);expect(resolved.charts[0]!.references[0]!.chartCacheIgnored).toBe(true)
 expect(Object.isFrozen(data.series[0]!.sizes)).toBe(true)
 workbook.sheets[0].cells[0].value.lexical='-1'
 const negative=await resolveNativePptxWorkbookCharts(inspection,async()=>JSON.stringify(workbook))
 expect(negative.charts).toEqual([]);expect(negative.refusals[0]!.reason).toMatch(/negative|size/)
})
it('refuses size/profile/paint/resource mismatches and noncanonical source order',async()=>{
 for(const mutate of [(s:any)=>delete s.bubbleScale,(s:any)=>s.bubbleScale=301,(s:any)=>s.sizeRepresents='height',(s:any)=>delete s.series[0].sizeReference,(s:any)=>s.series[0].sizeReference={...s.series[0].sizeReference,formula:'Lexical!A1:A2',range:parseChartWorkbookRange('Lexical!A1:A2')},(s:any)=>s.series[0].widthEmu=0,(s:any)=>s.series[0].sizes=['1'],(s:any)=>s.family='scatter',(s:any)=>s['']=1]){
  const {deck,result}=inputs();mutate(result.charts[0].source);await expect(decode(JSON.stringify(result),deck,sha)).rejects.toThrow()
 }
 const {deck,result}=inputs()
 await expect(decode(JSON.stringify(result).replace('"order":0','"order":-0'),deck,sha)).rejects.toThrow()
 await expect(decode(JSON.stringify({...result,'':1}),deck,sha)).rejects.toThrow()
 result.charts[0].workbook.sha256='b'.repeat(64);await expect(decode(JSON.stringify(result),deck,sha)).rejects.toThrow()
})
