import {expect,it} from 'vitest'
import {decodeNativePptxChartWorkbookInspection as decode,assertAdmittedChartWorkbookInspection} from './chartWorkbookInspection.js'
import {fixture,sha} from '../test/chartWorkbookFixture.js'
it('binds exact frame/chart/resource closure and returns only admitted immutable records',async()=>{
 const {deck,result}=fixture(),before=JSON.stringify(deck),got=await decode(JSON.stringify(result),deck,sha)
 expect(got).toEqual(result);expect(Object.isFrozen(got.charts[0]!.source.series[0]!.colors)).toBe(true)
 expect(()=>assertAdmittedChartWorkbookInspection(got)).not.toThrow();expect(()=>assertAdmittedChartWorkbookInspection(JSON.parse(JSON.stringify(got)))).toThrow(/decoder/)
 expect(JSON.stringify(deck)).toBe(before)
})
it('rejects identity drift, missing/extra resources, unreported charts, malformed families and caches masquerading as values',async()=>{
 const mutations=[
  (r:any)=>{r.charts[0].frame_sha256='b'.repeat(64)},(r:any)=>{r.charts[0].chart_sha256='b'.repeat(64)},(r:any)=>{r.charts[0].slide_sha256='b'.repeat(64)},
  (r:any)=>{r.charts[0].workbook.sha256='b'.repeat(64)},(r:any)=>{r.workbooks[0].bytesBase64='AQIE'},(r:any)=>{r.workbooks=[]},
  (r:any)=>{r.workbooks.push({...r.workbooks[0],part:'ppt/embeddings/unused.xlsx'})},(r:any)=>{r.charts=[];r.workbooks=[]},
  (r:any)=>{r.charts[0].source.series[0].values=['1','2']},(r:any)=>{r.charts[0].source.family='scatter'},
  (r:any)=>{r.charts[0].source.series[0].categoryReference.range.startRow=1},(r:any)=>{r.charts[0].source.plotVisibleOnly=true},
  (r:any)=>{r.charts[0].source.series[0].colors.pop()},(r:any)=>{r.charts[0].source.yAxis.crossAxisId=9},
 ]
 for(const change of mutations){const {deck,result}=fixture();change(result);await expect(decode(JSON.stringify(result),deck,sha)).rejects.toThrow()}
})
it('allows explicit omission coverage and refuses duplicate or unbound omissions',async()=>{
 const {deck,result}=fixture(),chart=result.charts[0]!
 const omitted={...result,charts:[],workbooks:[],omissions:[{slide_id:chart.slide_id,object_id:chart.object_id,reason:'Unqualified source.'}]}
 expect((await decode(JSON.stringify(omitted),deck,sha)).omissions).toHaveLength(1)
 await expect(decode(JSON.stringify({...omitted,omissions:[...omitted.omissions,...omitted.omissions]}),deck,sha)).rejects.toThrow()
 await expect(decode(JSON.stringify(result).replace('"protocol":','"protocol":"bad","protocol":'),deck,sha)).rejects.toThrow(/duplicate/)
})
