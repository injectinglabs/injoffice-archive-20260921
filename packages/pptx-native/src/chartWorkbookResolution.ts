import type {NativeLiteralBar,NativeLiteralConnected} from './types.js'
import type {NativePptxChartWorkbookInspection,NativePptxInspectedWorkbookChart} from './chartWorkbookInspectionTypes.js'
import type {ChartWorkbookReference,ChartWorkbookResolvedValues} from './chartWorkbookTypes.js'
import {assertAdmittedChartWorkbookInspection} from './chartWorkbookInspection.js'
import {assertChartWorkbookValues} from './chartWorkbookValueAdmission.js'

export interface NativeWorkbookBarData extends Omit<NativeLiteralBar,'profile'|'dataOrigin'>{profile:'workbook-bar-v1';dataOrigin:'embedded-workbook'}
export interface NativeWorkbookConnectedData extends Omit<NativeLiteralConnected,'profile'|'dataOrigin'>{profile:'workbook-line-v1'|'workbook-scatter-v1';dataOrigin:'embedded-workbook'}
export interface NativeResolvedWorkbookChart {
 readonly profile:'embedded-workbook-chart-v1'
 readonly packageSHA256:string;readonly sourceRevision:string
 readonly source:NativePptxInspectedWorkbookChart
 readonly data:NativeWorkbookBarData|NativeWorkbookConnectedData
 readonly references:readonly ChartWorkbookResolvedValues[]
}
const admitted=new WeakSet<NativeResolvedWorkbookChart>()
export function assertResolvedWorkbookChart(value:NativeResolvedWorkbookChart):void{if(!admitted.has(value))throw new TypeError('workbook chart must be assembled from admitted source and values')}
function freeze<T>(value:T):T{if(value&&typeof value==='object'){for(const child of Object.values(value))freeze(child);Object.freeze(value)}return value}

/** Combine separately admitted source paint/axes and authoritative worksheet
 * values. No source profile is relabeled as literal, and no deck is mutated. */
export function createResolvedWorkbookChart(inspection:NativePptxChartWorkbookInspection,chartIndex:number,values:readonly ChartWorkbookResolvedValues[]):NativeResolvedWorkbookChart {
 assertAdmittedChartWorkbookInspection(inspection)
 if(!Number.isInteger(chartIndex)||chartIndex<0||chartIndex>=inspection.charts.length||values.length<1||values.length>64)throw new RangeError('invalid workbook chart resolution budget')
 const source=inspection.charts[chartIndex]!,records=new Map<string,ChartWorkbookResolvedValues>(),used=new Set<string>()
 for(const value of values){
  assertChartWorkbookValues(value)
  if(value.workbookPart!==source.workbook.part||value.workbookSHA256!==source.workbook.sha256||value.workbookRevision!==`rev:${source.workbook.sha256}`)throw new TypeError('resolved workbook identity does not match chart source')
  const key=`${value.kind}\0${value.formula}\0${value.chartCacheIgnored}`;if(records.has(key))throw new TypeError('duplicate resolved chart reference');records.set(key,value)
 }
 const get=(reference:ChartWorkbookReference):string[]=>{
  const key=`${reference.kind}\0${reference.formula}\0${reference.cachePresent}`,value=records.get(key)
  if(!value||value.values.length!==reference.range.count||value.sheetName!==reference.range.sheet||value.chartCacheIgnored!==reference.cachePresent)throw new TypeError('resolved values do not match source reference')
  used.add(key);return [...value.values]
 }
 let categories:string[]|undefined
 const spec=source.source,series=spec.series.map(s=>{
  let title=s.title
  if(s.titleReference){title=get(s.titleReference)[0]!;if(title.length>1024)throw new RangeError('workbook series title budget exceeded')}
  if(s.categoryReference){const c=get(s.categoryReference);if(categories&&(categories.length!==c.length||categories.some((value,i)=>value!==c[i])))throw new TypeError('series category references resolve to different category values');categories=c}
  return {index:s.index,order:s.order,...(title===undefined?{}:{title}),values:get(s.valueReference),...(s.xReference?{xValues:get(s.xReference)}:{}),...(s.colors?{colors:[...s.colors]}:{}),...(s.color?{color:s.color,widthEmu:s.widthEmu!}:{})}
 })
 if(used.size!==records.size)throw new TypeError('unreferenced workbook values were supplied')
 const data:NativeWorkbookBarData|NativeWorkbookConnectedData=spec.family==='bar'?{
  profile:'workbook-bar-v1',dataOrigin:'embedded-workbook',barDirection:spec.barDirection==='col'?'column':'bar',grouping:'clustered',gapWidth:spec.gapWidth!,overlap:0,categories:categories!,series:series as NativeWorkbookBarData['series'],categoryAxis:spec.barDirection==='col'?spec.xAxis:spec.yAxis,valueAxis:spec.barDirection==='col'?spec.yAxis:spec.xAxis,
 }:{profile:spec.family==='line'?'workbook-line-v1':'workbook-scatter-v1',dataOrigin:'embedded-workbook',categories:categories??[],series:series as NativeWorkbookConnectedData['series'],xAxis:spec.xAxis,yAxis:spec.yAxis}
 const result=freeze({profile:'embedded-workbook-chart-v1' as const,packageSHA256:inspection.package_sha256,sourceRevision:inspection.source_revision,source,data,references:[...values]})
 admitted.add(result);return result
}
