import {expect,it} from 'vitest'
import {readFileSync} from 'node:fs'
import {decodeNativePptxChartWorkbookInspection,extractChartWorkbookReferences,createResolvedWorkbookChart} from '@injoffice/pptx-native'
import {fixture,sha} from '../../pptx-native/test/chartWorkbookFixture.js'
import {parseChartWorkbookRange} from '../../pptx-native/src/chartWorkbookRange.js'
import {createNativeWorkbookChartPaths} from './workbookChartPaths.js'
import {createNativeLiteralBarPaths} from './literalBar.js'
import {createNativeLiteralLinePaths} from './literalLine.js'

async function resolved(family:'bar'|'line'|'scatter'){
 const {deck,result}=fixture(),source=result.charts[0]!.source as any,s=source.series[0]
 source.family=family
 const reference=(formula:string,kind:'strRef'|'numRef')=>({kind,formula,range:parseChartWorkbookRange(formula),cachePresent:true})
 s.valueReference=reference('Lexical!A1:A2','numRef');s.categoryReference=reference('Lexical!C1:C2','strRef')
 if(family!=='bar'){delete source.barDirection;delete source.gapWidth;delete s.colors;s.color='#FF0000';s.widthEmu=1}
 if(family==='scatter'){delete s.categoryReference;s.xReference=s.valueReference;Object.assign(source.xAxis,{min:'-10',max:'10',crossesAt:'0'})}
 const inspection=await decodeNativePptxChartWorkbookInspection(JSON.stringify(result),deck,sha)
 const chart=inspection.charts[0]!
 // Complete trusted-callback contract stub. Actual dual-WASM source evidence is
 // separate; these tests exercise assembly, provenance and exact geometry.
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8'))
 workbook.source.package_sha256=`sha256:${chart.workbook.sha256}`;workbook.revision=`rev:${chart.workbook.sha256}`
 workbook.sheets[0].cells.push({row:1,column:0,ref:'A2',style_id:0,value:{kind:'number',storage:'number',lexical:'2',rich:false},editable:true},{row:1,column:2,ref:'C2',ooxml_type:'inlineStr',style_id:0,value:{kind:'string',storage:'inline',text:'World',rich:false},editable:true})
 const refs=[...new Map([s.categoryReference,s.xReference,s.valueReference].filter(Boolean).map(ref=>[JSON.stringify(ref),ref])).values()]
 const values=await extractChartWorkbookReferences(chart.workbook,refs,new Uint8Array([1,2,3]),async()=>JSON.stringify(workbook))
 return createResolvedWorkbookChart(inspection,0,values)
}
it('creates exact source-colored bars without converting workbook provenance to literal',async()=>{
 const chart=await resolved('bar'),paths=createNativeWorkbookChartPaths(chart,1000,100)
 expect(chart.data).toMatchObject({profile:'workbook-bar-v1',dataOrigin:'embedded-workbook',categories:['Hello','World']})
 expect(chart.data.series[0]!.values).toEqual(['001.2300','2'])
 expect(paths[0]).toMatchObject({color:'#FF0000',path:[{kind:'moveTo',x:150,y:44},{kind:'lineTo',x:350,y:44},{kind:'lineTo',x:350,y:50},{kind:'lineTo',x:150,y:50},{kind:'close'}]})
 expect(Object.isFrozen(chart.data.series[0]!.values)).toBe(true)
 expect(()=>createNativeWorkbookChartPaths(JSON.parse(JSON.stringify(chart)),1000,100)).toThrow(/admitted/)
 expect(()=>createNativeLiteralBarPaths(chart.data as any,1000,100)).toThrow(/literal bar profile/)
})
it('uses exact connected math for workbook line/XY while preserving family admission',async()=>{
 const line=await resolved('line'),scatter=await resolved('scatter')
 expect(createNativeWorkbookChartPaths(line,1000,100)[0]!.path).toEqual([{kind:'moveTo',x:250,y:44},{kind:'lineTo',x:750,y:40}])
 expect(createNativeWorkbookChartPaths(scatter,1000,100)[0]!.path).toEqual([{kind:'moveTo',x:562,y:44},{kind:'lineTo',x:600,y:40}])
 expect(()=>createNativeLiteralLinePaths(line.data as any,1000,100)).toThrow(/literal line profile/)
})
