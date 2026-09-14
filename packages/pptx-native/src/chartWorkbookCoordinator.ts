import {assertAdmittedChartWorkbookInspection} from './chartWorkbookInspection.js'
import {extractChartWorkbookContract,type ChartWorkbookExtractor} from './chartWorkbookExtract.js'
import {resolveChartWorkbookReferences} from './chartWorkbookResolve.js'
import {createResolvedWorkbookChart,type NativeResolvedWorkbookChart} from './chartWorkbookResolution.js'
import type {NativePptxChartWorkbookInspection,NativePptxInspectedWorkbookChart} from './chartWorkbookInspectionTypes.js'
import type {ChartWorkbookReference} from './chartWorkbookTypes.js'

export interface NativeWorkbookChartRefusal {readonly slideId:string;readonly objectId:string;readonly reason:string}
export interface NativeWorkbookChartResolution {readonly charts:readonly NativeResolvedWorkbookChart[];readonly refusals:readonly NativeWorkbookChartRefusal[]}
function chartReferences(chart:NativePptxInspectedWorkbookChart):readonly ChartWorkbookReference[]{
 const unique=new Map<string,ChartWorkbookReference>()
 for(const series of chart.source.series)for(const ref of [series.titleReference,series.categoryReference,series.xReference,series.valueReference,series.sizeReference])if(ref)unique.set(`${ref.kind}\0${ref.formula}\0${ref.cachePresent}`,ref)
 return [...unique.values()]
}
/** Explicit trusted XLSX injection. Each embedded resource is extracted once;
 * unsupported charts remain independent refusals and retain opaque fallback.
 * Across the operation: 16,384 referenced points and 2,000,000 scanned records. */
export async function resolveNativePptxWorkbookCharts(inspection:NativePptxChartWorkbookInspection,extract:ChartWorkbookExtractor):Promise<NativeWorkbookChartResolution>{
 assertAdmittedChartWorkbookInspection(inspection)
 const charts:NativeResolvedWorkbookChart[]=[],refusals:NativeWorkbookChartRefusal[]=inspection.omissions.map(o=>({slideId:o.slide_id,objectId:o.object_id,reason:o.reason}))
 let points=0,scans=0
 const refuse=(source:NativePptxInspectedWorkbookChart,error:unknown)=>{
  if(error instanceof Error&&error.name==='AbortError')throw error
  refusals.push({slideId:source.slide_id,objectId:source.object_id,reason:error instanceof Error?error.message.slice(0,2048):'Workbook source resolution unavailable'})
 }
 for(const resource of inspection.workbooks){
  const entries=inspection.charts.map((source,index)=>({source,index})).filter(e=>e.source.workbook.part===resource.part)
  try{
   const bytes=Uint8Array.from(atob(resource.bytesBase64),c=>c.charCodeAt(0))
   const workbook=await extractChartWorkbookContract(entries[0]!.source.workbook,bytes,extract)
   const records=workbook.sheets.reduce((sum,sheet)=>sum+sheet.cells.length+sheet.rows.length+sheet.columns.length,0)
   for(const {source,index} of entries){
    try{
     const refs=chartReferences(source),count=refs.reduce((sum,ref)=>sum+ref.range.count,0)
     if(points+count>16384||scans+records>2000000)throw new RangeError('aggregate workbook chart resolution budget exceeded')
     points+=count;scans+=records
     const values=resolveChartWorkbookReferences(source.workbook,refs,workbook)
     charts.push(createResolvedWorkbookChart(inspection,index,values))
    }catch(error){refuse(source,error)}
   }
  }catch(error){for(const {source} of entries)refuse(source,error)}
 }
 charts.sort((a,b)=>inspection.charts.indexOf(a.source)-inspection.charts.indexOf(b.source))
 return Object.freeze({charts:Object.freeze(charts),refusals:Object.freeze(refusals.map(r=>Object.freeze(r)))})
}
