import {assertResolvedWorkbookChart,type NativeResolvedWorkbookChart,type NativePptxDeck,type NativeElement} from '@injoffice/pptx-native'

/** Rebind an admitted outer resolution to the exact unchanged source deck being
 * compiled. The returned map owns the caller array before any async provider. */
export function bindWorkbookCharts(deck:NativePptxDeck,charts:readonly NativeResolvedWorkbookChart[]|undefined):ReadonlyMap<string,NativeResolvedWorkbookChart>{
 const result=new Map<string,NativeResolvedWorkbookChart>()
 if(charts===undefined)return result
 if(!Array.isArray(charts)||charts.length>64)throw new RangeError('workbook chart preview record budget exceeded')
 for(const chart of charts){
  assertResolvedWorkbookChart(chart)
  const source=chart.source,slide=deck.slides[source.slide_index]
  if(chart.sourceRevision!==deck.sourceRevision||chart.sourceRevision!==`rev-${chart.packageSHA256}`||!slide||source.slide_id!==slide.id||source.slide_part!==slide.source?.partName||source.slide_sha256!==slide.source?.fingerprintSha256)throw new TypeError('workbook chart does not belong to this source deck')
  let element:Extract<NativeElement,{kind:'chart'}>|undefined
  const walk=(elements:readonly NativeElement[])=>{for(const e of elements){if(e.kind==='group')walk(e.children);else if(e.kind==='chart'&&e.id===source.element_id)element=e}}
  walk(slide.elements)
  if(!element||source.object_id!==element.source?.objectId||source.frame_sha256!==element.source?.fingerprintSha256||source.chart_relationship_id!==element.chart.relationshipId||source.chart_part!==element.chart.chartPart||source.chart_sha256!==element.chart.opaqueRef.fingerprintSha256||result.has(element.id))throw new TypeError('workbook chart frame binding is stale or duplicated')
  result.set(element.id,chart)
 }
 return result
}
