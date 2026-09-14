import {expect,it} from 'vitest'
import type {NativeWorkbookV2} from '@injoffice/sheets/browser'
type NativeWorkbookUnsupportedV2=NativeWorkbookV2['unsupported'][number]
import {qualifyChartWorkbookDiagnostics as qualify} from './chartWorkbookDiagnostics.js'
const diagnostic=(change:Partial<NativeWorkbookUnsupportedV2>):NativeWorkbookUnsupportedV2=>({id:'source-id',code:'CELL_ATTRIBUTES',capability:'cell-markup',scope_id:'sheet:1',part_name:'xl/worksheets/sheet1.xml',preservation:'preserve-exact',message:'source markup',...change})
it('refuses intersecting cell/range source uncertainty but permits unrelated source cells',()=>{
 expect(()=>qualify([diagnostic({cell_ref:'B1'})],'1',['B1','B2'],false)).toThrow(/CELL_ATTRIBUTES/)
 expect(qualify([diagnostic({cell_ref:'C1'})],'1',['B1','B2'],false)).toEqual([])
 for(const range_ref of ['A1:C3','B1','B2:B3'])expect(()=>qualify([diagnostic({code:'FORMULA_GROUP_RANGE',range_ref})],'1',['B1','B2'],false)).toThrow(/FORMULA_GROUP_RANGE/)
 expect(qualify([diagnostic({range_ref:'C1:D8'})],'1',['B1','B2'],false)).toEqual([])
 for(const range_ref of ['B2:B1','D1:C2','XFE1:XFE2','A1:B1048577','A0:B1','A1:B2\n'])expect(()=>qualify([diagnostic({range_ref})],'1',['B1','B2'],false)).toThrow(/diagnostic range/)
})
it('preserves harmless styling/protection diagnostics but refuses unknown visibility metadata',()=>{
 const item=diagnostic({code:'STYLE_FONT_COLOR',capability:'styles',scope_id:'style:0'})
 const result=qualify([item],'1',['B1'],false)
 expect(result).toEqual([item]);expect(result[0]).not.toBe(item);expect(Object.isFrozen(result[0])).toBe(true)
 for(const code of ['COLUMN_DIMENSION_EXTRAS','ROW_EXTENSIONS','WORKSHEET_EXTENSIONS','SHEET_DECLARATION_ATTRIBUTES'])expect(()=>qualify([diagnostic({code})],'1',['B1'],false)).toThrow(code)
})
it('does not let shared-string-table uncertainty affect numeric-only source references',()=>{
 const item=diagnostic({code:'SHARED_STRING_TABLE_ATTRIBUTES',capability:'rich-text',scope_id:'workbook'})
 expect(qualify([item],'1',['B1'],false)).toHaveLength(1)
 expect(()=>qualify([item],'1',['A1'],true)).toThrow(/SHARED_STRING_TABLE_ATTRIBUTES/)
})

it('retains qualified metadata provenance without masking adjacent unknown records',()=>{
 const view=diagnostic({code:'WORKBOOK_VIEW_METADATA',capability:'workbook-features',scope_id:'workbook'})
 const dimension=diagnostic({code:'WORKSHEET_DIMENSION_METADATA',capability:'worksheet-features'})
 expect(qualify([view,dimension],'1',['B1'],false)).toEqual([view,dimension])
 for(const code of ['UNMODELED_WORKBOOK_FEATURE','UNMODELED_WORKSHEET_FEATURE']){
  const unknown=diagnostic({code,scope_id:code==='UNMODELED_WORKBOOK_FEATURE'?'workbook':'sheet:1'})
  for(const items of [[view,dimension,unknown],[unknown,view,dimension]])expect(()=>qualify(items,'1',['B1'],false)).toThrow(code)
 }
})
