import {expect,it} from 'vitest'
import type {NativeWorkbookV2} from '@injoffice/sheets/browser'
import {resolveChartWorkbookReferences as resolve} from './chartWorkbookResolve.js'
import {parseChartWorkbookRange} from './chartWorkbookRange.js'
import type {ChartWorkbookReference,ChartWorkbookBinding} from './chartWorkbookTypes.js'
const sha='a'.repeat(64),binding:ChartWorkbookBinding={relationshipId:'book',part:'ppt/embeddings/source.xlsx',sha256:sha,byteLength:100,autoUpdate:false}
const reference=(formula:string,kind:'numRef'|'strRef'='numRef'):ChartWorkbookReference=>({kind,formula,range:parseChartWorkbookRange(formula),cachePresent:true})
function fixture():NativeWorkbookV2{
 // Native XLSX engine identity; unrelated styles/capabilities are outside this
 // pure resolver test. Integration validates the complete extracted contract.
 return {protocol:'injoffice.xlsx.native',version:2,unsupported:[],revision:`rev:${sha}`,source:{authority:'exact-package-bytes',package_sha256:`sha256:${sha}`,workbook_part:'relocated/workbook.xml',dialect:'transitional'},sheets:[{id:'1',name:'Data',part_name:'relocated/worksheets/source.xml',state:'visible',rows:[{row:1,hidden:true}],columns:[{column:1,end_column:1,hidden:true}],sheet_format:{zero_height:false},cells:[
  {row:0,column:0,ref:'A1',value:{kind:'string',storage:'shared',text:'First',lexical:'0',rich:false}},
  {row:0,column:1,ref:'B1',value:{kind:'number',storage:'number',lexical:'-1.2500e+2',rich:false}},
  {row:0,column:2,ref:'C1',formula:{type:'normal',text:'1+1',cached:{kind:'number',storage:'number',lexical:'2',rich:false}}},
  {row:1,column:0,ref:'A2',value:{kind:'string',storage:'inline',text:'Second',rich:false}},
  {row:1,column:1,ref:'B2',value:{kind:'number',storage:'number',lexical:'+0.001',rich:false}},
 ]}]} as unknown as NativeWorkbookV2
}
it('joins exact source cells and retains lexemes, addresses and hidden metadata without cache use',()=>{
 const workbook=fixture(),before=JSON.stringify(workbook),refs=[reference('Data!$A$1:$A$2','strRef'),reference('Data!$B$1:$B$2')]
 const result=resolve(binding,refs,workbook)
 expect(result[0]!.values).toEqual(['First','Second']);expect(result[1]!.values).toEqual(['-1.2500e+2','+0.001'])
 expect(result[1]).toMatchObject({dataOrigin:'embedded-workbook',formula:'Data!$B$1:$B$2',workbookPart:binding.part,workbookSHA256:sha,workbookRevision:`rev:${sha}`,sheetId:'1',sheetName:'Data',sheetPart:'relocated/worksheets/source.xml',addresses:['B1','B2'],chartCacheIgnored:true})
 expect(result[1]!.visibility).toEqual([{sheetState:'visible',columnHidden:true,defaultRowsHidden:false},{sheetState:'visible',rowHidden:true,columnHidden:true,defaultRowsHidden:false}])
 expect(JSON.stringify(workbook)).toBe(before);expect(Object.isFrozen(result[1]!.values)).toBe(true)
 // The unreferenced C1 formula does not poison literal source references.
 expect(result).toHaveLength(2)
})
it('refuses referenced formulas even when their cached values look usable',()=>{
 for(const formula of [{type:'normal',text:'1+1',cached:{kind:'number',storage:'number',lexical:'2',rich:false}},{type:'shared',text:'',shared_index:0},{type:'array',text:'{1;2}',ref:'B1:B2'}]){
  const workbook=fixture(),cell=workbook.sheets[0]!.cells[1] as unknown as Record<string,unknown>
  delete cell.value;cell.formula=formula
  expect(()=>resolve(binding,[reference('Data!B1:B2')],workbook)).toThrow(/formula-backed/)
 }
})
it('refuses missing, numeric-text, error, date, rich and nonfinite source cells',()=>{
 const replacements=[undefined,{kind:'string',storage:'shared',text:'123',lexical:'0',rich:false},{kind:'error',storage:'error',lexical:'#N/A',rich:false},{kind:'date',storage:'date',lexical:'2026-01-01',rich:false},{kind:'number',storage:'number',lexical:'NaN',rich:false},{kind:'number',storage:'number',lexical:'1e101',rich:false}]
 for(const value of replacements){const workbook=fixture();(workbook.sheets[0]!.cells[1] as unknown as {value:unknown}).value=value;expect(()=>resolve(binding,[reference('Data!B1:B2')],workbook)).toThrow()}
 for(const value of [{kind:'string',storage:'shared',text:'rich',rich:true,runs:[{text:'rich'}]},{kind:'string',storage:'formula-string',text:'cached',rich:false}]){const workbook=fixture();(workbook.sheets[0]!.cells[0] as unknown as {value:unknown}).value=value;expect(()=>resolve(binding,[reference('Data!A1','strRef')],workbook)).toThrow()}
 expect(()=>resolve(binding,[reference('Data!B2:B3')],fixture())).toThrow(/missing/)
})
it('binds exact package revision, worksheet identity, formula coordinates and unique source cells',()=>{
 for(const change of [{sha256:'b'.repeat(64)},{sha256:sha+'\n'},{part:'../outside.xlsx'},{byteLength:8388609}])expect(()=>resolve({...binding,...change},[reference('Data!B1')],fixture())).toThrow()
 for(const mutate of [
  (w:any)=>{w.revision='rev:'+ 'b'.repeat(64)},
  (w:any)=>{w.source.authority='cache'},
  (w:any)=>{w.sheets.push(w.sheets[0])},
  (w:any)=>{w.sheets[0].cells.reverse()},
  (w:any)=>{w.sheets[0].cells[1].ref='A9'},
  (w:any)=>{w.sheets[0].columns.push({...w.sheets[0].columns[0]})},
 ]){const workbook=fixture();mutate(workbook);expect(()=>resolve(binding,[reference('Data!B1')],workbook)).toThrow()}
 const ref=reference('Data!B1');expect(()=>resolve(binding,[{...ref,range:{...ref.range,count:2}}],fixture())).toThrow(/coordinates/)
 expect(()=>resolve(binding,[reference('data!B1')],fixture())).toThrow(/missing/)
})
