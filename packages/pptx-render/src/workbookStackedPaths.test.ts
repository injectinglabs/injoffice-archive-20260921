import {it,expect} from 'vitest'
import {stackedWorkbookInputs} from '../../pptx-native/test/chartWorkbookStackedFixture.js'
import {sha} from '../../pptx-native/test/chartWorkbookFixture.js'
import {decodeNativePptxChartWorkbookInspection,resolveNativePptxWorkbookCharts} from '@injoffice/pptx-native'
import {createNativeWorkbookChartPaths} from './workbookChartPaths.js'
import {createNativeLiteralStackedBarPaths} from './literalStackedBar.js'
import {createNativeLiteralStackedLinePaths} from './literalStackedLine.js'
import {compileNativePptxSlide} from './compile.js'
const noTextLayout:import('./types.js').NativePptxTextLayout={
 manifest:{version:1,manifestId:'geometry-only',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused'}}],fallbackChains:[]},
 resolver:{providerId:'unused',providerRevision:'1',resolve(){throw Error('No text')},load(){throw Error('No font')}},shaper:{providerId:'unused',providerRevision:'1',shape(){throw Error('No glyphs')}},defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'},
}

it('renders exact signed workbook stacks in retained XML order with separate authority',async()=>{
 for(const family of ['bar','line'] as const)for(const grouping of ['stacked','percentStacked'] as const){
 const {deck,result,workbook}=stackedWorkbookInputs(family,grouping),inspection=await decodeNativePptxChartWorkbookInspection(JSON.stringify(result),deck,sha),resolved=await resolveNativePptxWorkbookCharts(inspection,async()=>JSON.stringify(workbook))
 expect(resolved.refusals).toEqual([]);const chart=resolved.charts[0]!,before=JSON.stringify(chart),vectors=createNativeWorkbookChartPaths(chart,600,600)
 expect(vectors.map(v=>v.seriesIndex)).toEqual([13,12,11,10]);expect(chart.data.dataOrigin).toBe('embedded-workbook')
 expect(()=>createNativeWorkbookChartPaths(structuredClone(chart),600,600)).toThrow(/admitted/)
 expect(()=>(family==='bar'?createNativeLiteralStackedBarPaths:createNativeLiteralStackedLinePaths)(chart.data as never,600,600)).toThrow(/literal/)
 deck.slides[0]!.elements=deck.slides[0]!.elements.filter(e=>e.kind==='chart')
 expect((await compileNativePptxSlide(deck,0,{textLayout:noTextLayout})).nodes[0]!.kind).toBe('image')
 const tree=await compileNativePptxSlide(deck,0,{textLayout:noTextLayout,workbookChartsPreview:[chart]}),node=tree.nodes[0]!
 expect(node.kind).toBe('group');if(family==='line')expect(node.clip?.kind).toBe('rect');expect(tree.diagnostics.some(d=>d.code==='chart.workbookDataPreview')).toBe(true)
 expect(JSON.stringify(chart)).toBe(before)
 if(grouping==='percentStacked'&&family==='line')expect(vectors[0]!.path).toEqual([{kind:'moveTo',x:150,y:330},{kind:'lineTo',x:450,y:300}])
 }
})
