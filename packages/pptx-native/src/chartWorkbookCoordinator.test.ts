import {expect,it} from 'vitest'
import {readFileSync} from 'node:fs'
import {fixture,sha} from '../test/chartWorkbookFixture.js'
import {parseChartWorkbookRange} from './chartWorkbookRange.js'
import {decodeNativePptxChartWorkbookInspection} from './chartWorkbookInspection.js'
import {resolveNativePptxWorkbookCharts} from './chartWorkbookCoordinator.js'

async function inputs(){
 const {deck,result}=fixture(),first=result.charts[0]!
 const reference=(formula:string,kind:'strRef'|'numRef')=>({kind,formula,range:parseChartWorkbookRange(formula),cachePresent:true})
 first.source.series[0]!.categoryReference=reference('Lexical!C1','strRef');first.source.series[0]!.valueReference=reference('Lexical!A1','numRef');first.source.series[0]!.colors=['#FF0000']
 const sourceElement=deck.slides[0]!.elements.find(e=>e.kind==='chart')!,clone=structuredClone(sourceElement)
 clone.id='second-chart';clone.source!.objectId='900';deck.slides[0]!.elements.push(clone)
 const second=structuredClone(first);second.element_id=clone.id;second.object_id='900';second.workbook.relationshipId='another-book';result.charts.push(second)
 const inspection=await decodeNativePptxChartWorkbookInspection(JSON.stringify(result),deck,sha)
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8'))
 workbook.source.package_sha256=`sha256:${first.workbook.sha256}`;workbook.revision=`rev:${first.workbook.sha256}`
 return {inspection,workbook,deck,result}
}
it('extracts each immutable resource once and resolves distinct chart relationship provenance',async()=>{
 const {inspection,workbook}=await inputs();let calls=0
 const result=await resolveNativePptxWorkbookCharts(inspection,async bytes=>{calls++;expect([...bytes]).toEqual([1,2,3]);bytes.fill(9);return JSON.stringify(workbook)})
 expect(calls).toBe(1);expect(result.refusals).toEqual([]);expect(result.charts).toHaveLength(2)
 expect(result.charts.map(c=>c.references[0]!.workbookRelationshipId)).toEqual(['book','another-book'])
 expect(result.charts[0]!.data.series[0]!.values).toEqual(['001.2300']);expect(inspection.workbooks[0]!.bytesBase64).toBe('AQID')
 expect(Object.isFrozen(result.charts)).toBe(true)
})
it('returns source-scoped refusals and propagates cancellation',async()=>{
 const {inspection}=await inputs()
 const result=await resolveNativePptxWorkbookCharts(inspection,async()=>'{"not":"a workbook"}')
 expect(result.charts).toEqual([]);expect(result.refusals.map(r=>r.objectId)).toEqual(inspection.charts.map(c=>c.object_id))
 await expect(resolveNativePptxWorkbookCharts(inspection,async()=>{throw new DOMException('Cancelled','AbortError')})).rejects.toMatchObject({name:'AbortError'})
 await expect(resolveNativePptxWorkbookCharts(structuredClone(inspection),async()=>'' )).rejects.toThrow(/source-bound decoder/)
})
