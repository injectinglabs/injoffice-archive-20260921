import {describe,it,expect} from 'vitest'
import {nativeTableFillPreview,nativeTableHeaderTextPreview} from './nativeTableFillPreview.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'

const revision=`sha256:${'a'.repeat(64)}`
const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:revision,charts:[],tables:[{part:'xl/tables/table1.xml',sheet_part:'xl/worksheets/sheet1.xml',name:'Example',ref:'B2:D6',style:'TableStyleMedium2',header_rows:1,total_rows:1,row_stripes:true,column_stripes:false,warnings:[],fill_preview:{header:'#156082',stripe:'#C0E6F5',body:'#FFFFFF',header_font_style_ids:[0],fill_style_ids:[0]}}]}
const fill={origin:'implicit-default' as const}
const paint=(row:number,column=1,input=objects)=>nativeTableFillPreview(input,revision,'xl/worksheets/sheet1.xml',row,column,fill,0)
describe('native table fill preview',()=>{
 it('bounds source-qualified style IDs cumulatively across table palettes',()=>{
  const ids=Array.from({length:4096},(_,i)=>i)
  const tables=Array.from({length:3},(_,i)=>({...objects.tables[0]!,part:`xl/tables/t${i}.xml`,fill_preview:{...objects.tables[0]!.fill_preview!,header_font_style_ids:ids,fill_style_ids:ids}}))
  expect(()=>decodeNativeWorkbookObjectsV1({...objects,tables:tables.slice(0,2)},revision)).not.toThrow()
  expect(()=>decodeNativeWorkbookObjectsV1({...objects,tables},revision)).toThrow()
 })
 it('uses white bold headers only for the matched table and default-font styles',()=>{
  expect(nativeTableHeaderTextPreview(objects,revision,'xl/worksheets/sheet1.xml',1,1,fill,0)).toBe(true)
  expect(nativeTableHeaderTextPreview(objects,revision,'xl/worksheets/sheet1.xml',1,1,fill,1)).toBe(false)
  expect(nativeTableHeaderTextPreview(objects,revision,'xl/worksheets/sheet1.xml',2,1,fill,0)).toBe(false)
  const adjacent={...objects,tables:[...objects.tables,{...objects.tables[0]!,part:'xl/tables/second.xml',ref:'F2:H6',fill_preview:{...objects.tables[0]!.fill_preview!,header_font_style_ids:[7]}}]}
  expect(nativeTableHeaderTextPreview(adjacent,revision,'xl/worksheets/sheet1.xml',1,1,fill,7)).toBe(false)
 })
 it('uses source-bound header and alternating body fills, never totals or out-of-range cells',()=>{
  expect(paint(1)).toBe('#156082');expect(paint(2)).toBe('#C0E6F5');expect(paint(3)).toBe('#FFFFFF');expect(paint(4)).toBe('#C0E6F5');expect(paint(5)).toBeUndefined();expect(paint(1,0)).toBeUndefined()
 })
 it('does not replace direct fills, absent provenance, stale results or ambiguous overlapping tables',()=>{
  expect(nativeTableFillPreview(objects,revision,'xl/worksheets/sheet1.xml',1,1,fill,9)).toBeUndefined()
  for(const provenance of [undefined,{origin:'styles-record' as const,fill_id:2,color:'#FFFFFF'}])expect(nativeTableFillPreview(objects,revision,'xl/worksheets/sheet1.xml',1,1,provenance,0)).toBeUndefined()
  expect(nativeTableFillPreview(objects,'stale','xl/worksheets/sheet1.xml',1,1,fill,0)).toBeUndefined()
  expect(paint(1,1,{...objects,tables:[...objects.tables,...objects.tables]})).toBeUndefined()
 })
 it('validates optional fill schema without accepting CSS or unqualified styles',()=>{
  expect(decodeNativeWorkbookObjectsV1(objects,revision)).toEqual(objects)
  for(const mutate of [(v:any)=>v.tables[0].fill_preview.header='url(https://invalid)',(v:any)=>v.tables[0].style='Unqualified']){const value=structuredClone(objects);mutate(value);expect(()=>decodeNativeWorkbookObjectsV1(value,revision)).toThrow()}
 })
})
