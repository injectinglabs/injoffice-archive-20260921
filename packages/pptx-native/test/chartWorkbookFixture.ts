import {readFileSync} from 'node:fs'
import type {NativePptxDeck} from '@injoffice/pptx-native'
import {parseChartWorkbookRange} from '../src/chartWorkbookRange.js'
export const sha='a'.repeat(64),workbookSHA='039058c6f2c0cb492c533b0a4d14ef77cc0f78abccced5287d84a1a2011cfb81'
export function fixture(){
 const deck=JSON.parse(readFileSync(new URL('../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json',import.meta.url),'utf8')) as NativePptxDeck
 deck.sourceRevision=`rev-${sha}`
 const slide=deck.slides[0]!,chart=slide.elements.find(e=>e.kind==='chart')!;if(chart.kind!=='chart')throw Error('chart')
 const reference=(formula:string,kind:'strRef'|'numRef')=>({kind,formula,range:parseChartWorkbookRange(formula),cachePresent:true})
 const source={family:'bar',barDirection:'col',gapWidth:150,plotVisibleOnly:false,xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:true},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:true,min:'-10',max:'10',crossesAt:'0'},series:[{index:0,order:0,categoryReference:reference('Data!A1:A2','strRef'),valueReference:reference('Data!B1:B2','numRef'),colors:['#FF0000','#0000FF']}]}
 const resource={part:'ppt/embeddings/source.xlsx',sha256:workbookSHA,byteLength:3,bytesBase64:'AQID'}
 const result={protocol:'pptx-chart-workbook-inspection-v1',package_sha256:sha,source_revision:`rev-${sha}`,charts:[{slide_id:slide.id,slide_index:0,slide_part:slide.source!.partName,slide_sha256:slide.source!.fingerprintSha256,element_id:chart.id,object_id:chart.source!.objectId,frame_sha256:chart.source!.fingerprintSha256,chart_relationship_id:chart.chart.relationshipId,chart_part:chart.chart.chartPart,chart_sha256:chart.chart.opaqueRef.fingerprintSha256,source,workbook:{relationshipId:'book',part:resource.part,sha256:resource.sha256,byteLength:3}}],workbooks:[resource],omissions:[]}
 return {deck,result}
}
