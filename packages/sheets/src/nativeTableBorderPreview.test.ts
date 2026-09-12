import {it,expect} from 'vitest'
import {nativeTableBorderPreview} from './nativeTableBorderPreview.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
const revision=`sha256:${'a'.repeat(64)}`
const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:revision,charts:[],tables:[{part:'xl/tables/t.xml',sheet_part:'xl/worksheets/s.xml',name:'Original',ref:'A1:C4',style:'TableStyleMedium2',header_rows:1,total_rows:1,row_stripes:true,column_stripes:false,warnings:[],fill_preview:{header:'#156082',stripe:'#C0E6F5',body:'#FFFFFF',fill_style_ids:[0],header_font_style_ids:[0]},border_preview:{color:'#44B3E1',totals_color:'#156082',width_points:1,totals_width_points:3,style_ids:[0]}}]}
const neighbors={top:0,right:0,bottom:0,left:0}
it('qualifies measured horizontal/outer edges and totals divider, without interior vertical lines',()=>{
 expect(nativeTableBorderPreview(objects,revision,'xl/worksheets/s.xml',0,0,0,{...neighbors,top:null,left:null})).toEqual({top:{style:'solid',color:'#44B3E1',widthPoints:1},left:{style:'solid',color:'#44B3E1',widthPoints:1},right:{style:'none',color:'#44B3E1',widthPoints:0},bottom:{style:'solid',color:'#44B3E1',widthPoints:1}})
 expect(nativeTableBorderPreview(objects,revision,'xl/worksheets/s.xml',3,1,0,neighbors)?.top).toEqual({style:'double',color:'#156082',widthPoints:3})
 expect(nativeTableBorderPreview(objects,revision,'xl/worksheets/s.xml',2,1,0,neighbors)?.bottom).toEqual({style:'double',color:'#156082',widthPoints:3})
})
it('never overrides adjacent explicit, unknown or conflicting table borders',()=>{
 expect(nativeTableBorderPreview(objects,revision,'xl/worksheets/s.xml',1,1,0,{...neighbors,right:1,top:undefined,left:null})).toEqual({bottom:{style:'solid',color:'#44B3E1',widthPoints:1}})
 expect(nativeTableBorderPreview(objects,revision,'xl/worksheets/s.xml',1,1,9,neighbors)).toBeUndefined()
 const adjacent={...objects,tables:[...objects.tables,{...objects.tables[0]!,part:'xl/tables/u.xml',ref:'D1:F4'}]}
 expect(nativeTableBorderPreview(adjacent,revision,'xl/worksheets/s.xml',1,2,0,neighbors)?.right).toBeUndefined()
 expect(nativeTableBorderPreview(objects,'stale','xl/worksheets/s.xml',1,1,0,neighbors)).toBeUndefined()
})
it('validates exact bounded border projection fields',()=>{
 expect(decodeNativeWorkbookObjectsV1(objects,revision)).toEqual(objects)
 const value=structuredClone(objects) as any;value.tables[0].border_preview.width_points=99
 expect(()=>decodeNativeWorkbookObjectsV1(value,revision)).toThrow()
})
it('shares the cumulative style budget across borders and differential formats',()=>{
 const ids=Array.from({length:4096},(_,i)=>i)
 const tables=Array.from({length:3},(_,i)=>({...objects.tables[0]!,part:`xl/tables/t${i}.xml`,fill_preview:{...objects.tables[0]!.fill_preview!,fill_style_ids:[],header_font_style_ids:[]},border_preview:{...objects.tables[0]!.border_preview!,style_ids:ids},number_formats:[{ref:'A1:C4',dxf_id:0,number_format:'0.00',style_ids:ids}]}))
 expect(()=>decodeNativeWorkbookObjectsV1({...objects,tables:tables.slice(0,2)},revision)).not.toThrow()
 expect(()=>decodeNativeWorkbookObjectsV1({...objects,tables},revision)).toThrow()
})
