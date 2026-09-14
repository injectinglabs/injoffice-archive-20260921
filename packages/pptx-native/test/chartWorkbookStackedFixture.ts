import {readFileSync} from 'node:fs'
import {fixture} from './chartWorkbookFixture.js'
import {parseChartWorkbookRange} from '../src/chartWorkbookRange.js'
export function stackedWorkbookInputs(family:'bar'|'line'='bar',grouping:'stacked'|'percentStacked'='stacked'){
 const {deck,result:base}=fixture(),result:any=base,source=result.charts[0].source
 source.family=family;source.grouping=grouping
 if(family==='bar')source.overlap=100;else{delete source.barDirection;delete source.gapWidth}
 source.yAxis.min=grouping==='stacked'?'-10':'-1';source.yAxis.max=grouping==='stacked'?'10':'1'
 const ref=(formula:string,kind:'numRef'|'strRef')=>({kind,formula,range:parseChartWorkbookRange(formula),cachePresent:true})
 source.series=[4,-2,3,-1].map((_,i)=>({index:10+i,order:i,categoryReference:ref('Lexical!E1:E2','strRef'),valueReference:ref(`Lexical!${'ABCD'[i]}1:${'ABCD'[i]}2`,'numRef'),...(family==='bar'?{colors:['#123456','#ABCDEF']}:{color:'#123456',widthEmu:12700})})).reverse()
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8'))
 workbook.source.package_sha256=`sha256:${result.charts[0].workbook.sha256}`;workbook.revision=`rev:${result.charts[0].workbook.sha256}`
 workbook.sheets[0].cells=[];workbook.unsupported=[]
 for(let row=0;row<2;row++)for(let column=0;column<5;column++)workbook.sheets[0].cells.push({row,column,ref:`${'ABCDE'[column]}${row+1}`,style_id:0,...(column===4?{ooxml_type:'inlineStr',value:{kind:'string',storage:'inline',text:row?'B':'A',rich:false}}:{value:{kind:'number',storage:'number',lexical:row?'0':['4','-2','3','-1'][column],rich:false}}),editable:true})
 return {deck,result,workbook}
}
