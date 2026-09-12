import {it,expect} from 'vitest'
import {nativeTableNumberFormatPreview} from './nativeTableNumberFormatPreview.js'
import type {NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
const revision=`sha256:${'a'.repeat(64)}`
function objects(format:string):NativeWorkbookObjectsV1{return {protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:revision,charts:[],tables:[{part:'xl/tables/t.xml',sheet_part:'xl/worksheets/s.xml',name:'Original',ref:'A1:C4',style:'TableStyleMedium2',header_rows:1,total_rows:1,row_stripes:true,column_stripes:false,warnings:[],number_formats:[{ref:'C4:C4',dxf_id:0,number_format:format,style_ids:[0]}]}]}}
const accounting='_-* #,##0.00\\ "Ft"_-;\\-* #,##0.00\\ "Ft"_-;_-* "-"??\\ "Ft"_-;_-@_-'
const preview=(lexical:string,format=accounting)=>nativeTableNumberFormatPreview(objects(format),revision,'xl/worksheets/s.xml',3,2,0,'number',lexical)
it('retains exact saved decimals and literal currency with explicit accounting-layout limitation',()=>{
 expect(preview('2.75')?.text).toBe('2.75 Ft');expect(preview('-1234.567')?.text).toBe('-1,234.57 Ft');expect(preview('0')?.text).toBe('- Ft');expect(preview('2.75')?.warning).toMatch(/padding/)
 expect(preview('9007199254740993','0.00" X"')?.text).toBe('9007199254740993.00 X')
 expect(preview('2.75','0.00" X"')?.warning).toBeUndefined()
})
it('does not broaden strict formats into unknown conditional/scaled/date semantics',()=>{
 for(const format of ['[Red]0.00','0.00E+00','0,','0.00"x"0','[>=1]0.00;0.0','"unterminated','0.0*'])expect(preview('2.75',format)?.text).toBeUndefined()
})
it('joins source revision, worksheet, range and eligible cell style',()=>{
 const value=objects('0.00')
 expect(nativeTableNumberFormatPreview(value,'stale','xl/worksheets/s.xml',3,2,0,'number','2.75')).toBeUndefined()
 expect(nativeTableNumberFormatPreview(value,revision,'xl/worksheets/s.xml',3,1,0,'number','2.75')).toBeUndefined()
 expect(nativeTableNumberFormatPreview(value,revision,'xl/worksheets/s.xml',3,2,9,'number','2.75')).toBeUndefined()
 value.tables.push({...value.tables[0]!,part:'xl/tables/other.xml'})
 expect(nativeTableNumberFormatPreview(value,revision,'xl/worksheets/s.xml',3,2,0,'number','2.75')).toBeUndefined()
})
