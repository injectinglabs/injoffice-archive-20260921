import {describe,expect,it} from 'vitest'
import {isNativePreviewPartPathV1} from './nativePreviewPartPathV1.js'
import {decodeNativeConditionalFillPreviewsV1} from './nativeConditionalFillPreviewV1.js'
import {decodeNativeSheetPageSettingsV1} from './nativeSheetPageSettingsV1.js'
import {decodeNativeSheetPrintAreasV1} from './nativeSheetPrintAreasV1.js'
import {decodeNativeSheetPrintTitlesV1} from './nativeSheetPrintTitlesV1.js'
import {decodeNativeDrawingObjectsV1} from './nativeDrawingObjectsV1.js'
import {decodeNativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'

const decoders=[decodeNativeConditionalFillPreviewsV1,decodeNativeSheetPageSettingsV1,decodeNativeSheetPrintAreasV1,decodeNativeSheetPrintTitlesV1]
const entry=(part:string)=>({sheet_id:'7',sheet_part:part,status:'unavailable',warnings:['Read-only source preview unavailable']})
const safe=['Sheets/预算.xml','Sheets/年度 预算.xml','绘图/图表 ①.xml','Sheets/café.xml','Sheets/cafe\u0301.xml','Sheets/📊.xml','Sheets/%2e%2e.xml','目录/%2F.xml','a'.repeat(1024)]
const unsafe=['','/Sheets/预算.xml','Sheets\\预算.xml','Sheets//预算.xml','Sheets/./预算.xml','Sheets/../预算.xml','Sheets/预算.xml/','a'.repeat(1025),...Array.from({length:32},(_,i)=>`Sheets/a${String.fromCharCode(i)}.xml`),'Sheets/a\u007f.xml']

describe('exact native supplemental preview part identities',()=>{
 it('accepts Unicode and spaces in every supplemental decoder without transforming identity',()=>{
  for(const part of safe){
   expect(isNativePreviewPartPathV1(part)).toBe(true)
   for(const decode of decoders)expect(decode([entry(part)])[0]!.sheet_part).toBe(part)
   const drawing={sheet_id:'7',sheet_part:part,drawing_part:part,ordinal:1,kind:'unsupported',warnings:['Source drawing']}
   expect(decodeNativeDrawingObjectsV1([drawing])[0]).toEqual(drawing)
  }
 })
 it('rejects controls, traversal segments, separators, empty paths and excessive length in every decoder',()=>{
  for(const part of unsafe){
   expect(isNativePreviewPartPathV1(part)).toBe(false)
   for(const decode of decoders)expect(()=>decode([entry(part)])).toThrow()
   expect(()=>decodeNativeDrawingObjectsV1([{sheet_id:'7',sheet_part:part,drawing_part:'',ordinal:0,kind:'unsupported',warnings:['Source drawing']}])).toThrow()
  }
  for(const value of [null,undefined,12,{},[],new String('Sheets/预算.xml')])expect(isNativePreviewPartPathV1(value)).toBe(false)
 })
 it('retains only the existing explicitly empty drawing-part case',()=>{
  const drawing={sheet_id:'7',sheet_part:'Sheets/预算.xml',drawing_part:'',ordinal:0,kind:'unsupported',warnings:['Unresolved source drawing']}
  expect(decodeNativeDrawingObjectsV1([drawing])[0]).toEqual(drawing)
  expect(()=>decodeNativeDrawingObjectsV1([{...drawing,kind:'chart',chart_part:'图表/图.xml'}])).toThrow()
  const chart={...drawing,kind:'chart',drawing_part:'绘图/绘图.xml',chart_part:'图表/图.xml',ordinal:1,anchor:{kind:'oneCellAnchor',from:{row:0,column:0,row_offset_emu:0,column_offset_emu:0},width_emu:1,height_emu:1}}
  expect(decodeNativeDrawingObjectsV1([chart])[0]!.chart_part).toBe(chart.chart_part)
  for(const part of unsafe){
   expect(()=>decodeNativeDrawingObjectsV1([{...chart,chart_part:part}])).toThrow()
   if(part!=='')expect(()=>decodeNativeDrawingObjectsV1([{...drawing,drawing_part:part}])).toThrow()
  }
 })
 it('decodes a complete public supplemental envelope without normalizing or merging distinct Unicode spellings',()=>{
  const parts=['Sheets/café.xml','Sheets/cafe\u0301.xml']
  const entries=parts.map((part,i)=>({...entry(part),sheet_id:String(i+7)}))
  const object={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:`sha256:${'a'.repeat(64)}`,tables:[],charts:[],page_settings:entries,print_areas:entries,print_titles:entries,conditional_fills:entries}
  const decoded=decodeNativeWorkbookObjectsV1(object,object.package_sha256)
  expect(decoded.page_settings!.map(e=>e.sheet_part)).toEqual(parts)
  expect(decoded.conditional_fills!.map(e=>e.sheet_part)).toEqual(parts)
 })
})
