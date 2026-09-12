import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {decodeNativeWorkbook,type NativeCell} from './nativeRoundTrip'
import {nativeSheetPageCellPreview as preview} from './nativeSheetPageCellPreview'
import type {NativeWorkbookObjectsV1} from '@injoffice/sheets/browser'
const source=decodeNativeWorkbook(JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json',import.meta.url),'utf8')))
function fixture(format='General'){
 const workbook={...source,styles:[{...source.styles[0]!,effective:{...source.styles[0]!.effective,number_format:format,unsupported:[]}}]}
 const cell:NativeCell={row:0,column:0,ref:'A1',style_id:0,editable:true,value:{kind:'number',storage:'number',lexical:'5.3333333333333304',rich:false}}
 const objects={protocol:'injoffice.xlsx.preview-objects' as const,version:1 as const,package_sha256:workbook.source.package_sha256,tables:[],charts:[]}
 return {workbook,cell,objects}
}
describe('page-only compact display',()=>{
 it('defaults off, applies only by explicit choice, and preserves saved formulas/source',()=>{
  const {workbook,cell,objects}=fixture(),before=JSON.stringify({workbook,cell,objects})
  expect(preview(workbook,cell,objects,'s.xml').text).toBe('5.3333333333333304')
  expect(preview(workbook,cell,objects,'s.xml',true)).toMatchObject({text:'5.333333',compacted:true,cached:false})
  const formula={...cell,value:undefined,formula:{type:'normal' as const,text:'1/3',cached:cell.value}}
  expect(preview(workbook,formula,objects,'s.xml',true)).toMatchObject({text:'5.333333',cached:true})
  expect(preview(workbook,{...formula,formula:{type:'normal',text:'1/3'}},objects,'s.xml',true)).toMatchObject({text:'—',cached:false,warnings:[expect.stringContaining('No saved formula')]})
  expect(JSON.stringify({workbook,cell,objects})).toBe(before)
 })
 it('does not replace authored number, date, accounting, string or error formats',()=>{
  for(const format of ['0.00','0.00%','yyyy-mm-dd','_-* #,##0.00\\ "Ft"_-;\\-* #,##0.00\\ "Ft"_-']){
   const {workbook,cell,objects}=fixture(format)
   expect(preview(workbook,cell,objects,'s.xml',true)).toEqual(preview(workbook,cell,objects,'s.xml',false))
  }
  const {workbook,cell,objects}=fixture()
  for(const value of [{kind:'string',storage:'inline-string',text:'5.3333333333333304',rich:false},{kind:'error',storage:'error',lexical:'#VALUE!',rich:false}]as const){
   expect(preview(workbook,{...cell,value}as NativeCell,objects,'s.xml',true)).toMatchObject({compacted:false})
  }
 })
 it('reports malformed compact values, unavailable styles and text truncation',()=>{
  const {workbook,cell,objects}=fixture()
  expect(preview(workbook,{...cell,value:{...cell.value!,lexical:'1e129'}},objects,'s.xml',true).warnings.join(' ')).toContain('Compact number preview unavailable')
  expect(preview({...workbook,styles:[]},cell,objects,'s.xml',true)).toMatchObject({compacted:false,warnings:[expect.stringContaining('Number format unavailable')]})
  expect(preview(workbook,{...cell,value:{kind:'string',storage:'inline-string',text:'x'.repeat(2049),rich:false}}as NativeCell,objects,'s.xml',true).truncated).toBe(true)
 })
 it('does not replace source-joined table number format overrides',()=>{
  const {workbook,cell,objects}=fixture()
  const tableObjects:NativeWorkbookObjectsV1={...objects,tables:[{part:'xl/tables/t.xml',sheet_part:'s.xml',name:'Original',ref:'A1:A3',style:'TableStyleMedium2',header_rows:1,total_rows:0,row_stripes:true,column_stripes:false,warnings:[],number_formats:[{ref:'A1:A1',dxf_id:0,number_format:'0.00" X"',style_ids:[0]}]}]}
  expect(preview(workbook,cell,tableObjects,'s.xml',true)).toEqual(preview(workbook,cell,tableObjects,'s.xml',false))
  expect(preview(workbook,cell,tableObjects,'s.xml',true)).toMatchObject({text:'5.33 X',compacted:false})
 })
})
