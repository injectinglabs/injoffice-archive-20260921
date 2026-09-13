import {createRequire} from 'node:module'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {decodeNativeSheetPrintAreaSetsV1,decodeNativeWorkbookObjectsV1,selectNativeSheetPrintAreaSetV1,selectNativeSheetPrintAreaV1,compileNativeSheetPrintAreaSetPreviewV1,compileNativeSheetPagePreviewV1,projectNativeWorkbookV2,compileNativeSheetGeometryV2,createNativeMaximumDigitWidthAuthorityV2,type NativeWorkbookV2,type NativeWorkbookObjectsV1,type NativeSheetViewportV2} from './index.js'
const require=createRequire(import.meta.url)
const ranges=[{row:8,column:5,end_row:10,end_column:7},{row:1,column:1,end_row:3,end_column:3}]
function fixture(areas:NativeSheetViewportV2[]=ranges){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 const model=projectNativeWorkbookV2(workbook),font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
 const part=model.sheets[0]!.mutation_authority.source_part
 const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:model.source.package_sha256,tables:[],charts:[],print_area_sets:[{sheet_id:'7',sheet_part:part,status:'available',areas:structuredClone(areas),warnings:['Source ordered areas']}],page_settings:[{sheet_id:'7',sheet_part:part,status:'available',warnings:['Source settings'],settings:{paper:'Letter',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}}]}
 const compile=(area:NativeSheetViewportV2)=>compileNativeSheetGeometryV2(model,'7',area,createNativeMaximumDigitWidthAuthorityV2(model,font))
 return {model,objects,compile,part}
}
describe('source-bound multi-range print areas',()=>{
 it('decodes owned plain data and preserves source order rather than sorting or bounding-union',()=>{
  const {objects}=fixture(),copy=decodeNativeWorkbookObjectsV1(objects,objects.package_sha256)
  expect(copy).toEqual(objects)
  copy.print_area_sets![0]!.areas![0]={...copy.print_area_sets![0]!.areas![0]!,row:9}
  expect(objects.print_area_sets![0]!.areas).toEqual(ranges)
  delete objects.print_area_sets
  expect(Object.hasOwn(decodeNativeWorkbookObjectsV1(objects,objects.package_sha256),'print_area_sets')).toBe(false)
  expect(()=>decodeNativeWorkbookObjectsV1({...objects,print_area_sets:undefined},objects.package_sha256)).toThrow()
 })
 it('rejects overlaps, duplicate identities, empty/oversized sets and closed-shape violations',()=>{
  const entry=fixture().objects.print_area_sets![0]!
  for(const input of [[entry,entry],[{...entry,areas:[]}],[{...entry,areas:Array(17).fill(ranges[0])}],[{...entry,areas:[ranges[0],ranges[0]]}],[{...entry,areas:[ranges[0],{row:10,column:7,end_row:12,end_column:9}]}],[{...entry,status:'unavailable'}],[{...entry,extra:1}],[{...entry,areas:[{...ranges[0],extra:1}]}],[Object.assign(Object.create({extra:true}),entry)],[{...entry,warnings:['bad\nwarning']}],[{...entry,areas:[{row:-0,column:0,end_row:1,end_column:1}]}]])expect(()=>decodeNativeSheetPrintAreaSetsV1(input)).toThrow()
  let called=false;const accessor={...entry};Object.defineProperty(accessor,'areas',{enumerable:true,get(){called=true;return ranges}})
  expect(()=>decodeNativeSheetPrintAreaSetsV1([accessor])).toThrow();expect(called).toBe(false)
  expect(decodeNativeSheetPrintAreaSetsV1([{...entry,areas:[{row:0,column:0,end_row:0,end_column:0},{row:0,column:1,end_row:0,end_column:1}]}])[0]!.areas).toHaveLength(2)
 })
 it('joins frozen selections to package, sheet and source part',()=>{
  const {model,objects}=fixture(),areas=selectNativeSheetPrintAreaSetV1(model,'7',objects)
  expect(areas).toEqual(ranges);expect(Object.isFrozen(areas)).toBe(true);expect(areas.every(Object.isFrozen)).toBe(true)
  expect(()=>selectNativeSheetPrintAreaSetV1({...model},'7',objects)).toThrow('projected')
  expect(()=>selectNativeSheetPrintAreaSetV1(model,'8',objects)).toThrow('join')
  expect(()=>selectNativeSheetPrintAreaSetV1(model,'7',{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
  objects.print_area_sets![0]!.sheet_part='wrong.xml'
  expect(()=>selectNativeSheetPrintAreaSetV1(model,'7',objects)).toThrow('join')
 })
 it('accepts the complete bounded 64-sheet inventory with 16 disjoint areas per sheet',()=>{
  const {objects}=fixture(),entry=objects.print_area_sets![0]!
  const entries=Array.from({length:64},(_,i)=>({...entry,sheet_id:String(i+1),sheet_part:`Worksheets/Sheet${i+1}.xml`,areas:Array.from({length:16},(_,column)=>({row:0,column,end_row:0,end_column:column}))}))
  expect(decodeNativeSheetPrintAreaSetsV1(entries)).toEqual(entries)
  expect(decodeNativeWorkbookObjectsV1({...objects,print_area_sets:entries},objects.package_sha256).print_area_sets).toEqual(entries)
 })
 it('uses legacy single areas only when additive metadata is absent; existing selector remains unchanged',()=>{
  const {model,objects,part,compile}=fixture()
  objects.print_areas=[{sheet_id:'7',sheet_part:part,status:'available',area:ranges[1]!,warnings:['Legacy single']}]
  expect(selectNativeSheetPrintAreaV1(model,'7',objects)).toEqual(ranges[1])
  expect(selectNativeSheetPrintAreaSetV1(model,'7',objects)).toEqual(ranges)
  objects.print_area_sets=[{sheet_id:'7',sheet_part:part,status:'unavailable',warnings:['Unsupported source set']}]
  expect(()=>selectNativeSheetPrintAreaSetV1(model,'7',objects)).toThrow('unavailable')
  objects.print_area_sets=[];expect(()=>selectNativeSheetPrintAreaSetV1(model,'7',objects)).toThrow('join')
  delete objects.print_area_sets;expect(selectNativeSheetPrintAreaSetV1(model,'7',objects)).toEqual([ranges[1]])
  expect(compileNativeSheetPrintAreaSetPreviewV1([compile(ranges[1]!)],objects).total_pages).toBe(1)
  delete objects.print_areas;expect(()=>selectNativeSheetPrintAreaSetV1(model,'7',objects)).toThrow('join')
 })
 it('enforces per-area geometry limits and the aggregate cell budget without truncation',()=>{
  for(const areas of [[{row:0,column:0,end_row:4096,end_column:0}],[{row:0,column:0,end_row:0,end_column:1024}],[{row:0,column:0,end_row:999,end_column:59},{row:1000,column:0,end_row:1999,end_column:59}]]){
   const {model,objects}=fixture(areas)
   expect(()=>decodeNativeSheetPrintAreaSetsV1(objects.print_area_sets)).not.toThrow()
   expect(()=>selectNativeSheetPrintAreaSetV1(model,'7',objects)).toThrow('viewport limits')
  }
 })
 it('plans every non-A1 area independently in authored order and rejects missing, reordered or unbranded geometry',()=>{
  const {objects,compile}=fixture(),geometries=ranges.map(compile),before=JSON.stringify(objects)
  const result=compileNativeSheetPrintAreaSetPreviewV1(geometries,objects)
  expect(result.protocol).toBe('injoffice.xlsx.print-area-set-pages');expect(result.total_pages).toBe(2)
  expect(result.areas.map(a=>a.viewport)).toEqual(ranges);expect(result.areas.map(a=>a.area_index)).toEqual([0,1])
  expect(result.areas.map(a=>a.plan)).toEqual(geometries.map(g=>compileNativeSheetPagePreviewV1(g,objects)))
  expect(result.areas.map(a=>a.plan.pages[0]!.rows.start)).toEqual([8,1]);expect(JSON.stringify(objects)).toBe(before)
  for(const bad of [[],[geometries[0]!],[geometries[0]!,geometries[0]!],[...geometries].reverse(),[{...geometries[0]!},geometries[1]!],[compile({row:0,column:0,end_row:10,end_column:7}),geometries[1]!]])expect(()=>compileNativeSheetPrintAreaSetPreviewV1(bad,objects)).toThrow()
  objects.print_area_sets![0]!.sheet_part='wrong.xml';expect(()=>compileNativeSheetPrintAreaSetPreviewV1(geometries,objects)).toThrow('join')
 })
 it('refuses an impossible later fit without returning the earlier successful area',()=>{
  const selected=[{row:0,column:0,end_row:0,end_column:0},{row:1,column:0,end_row:2,end_column:999}]
  const {objects,compile}=fixture(selected),geometries=selected.map(compile)
  Object.assign(objects.page_settings![0]!.settings!,{fit_to_page:{width:1,height:1}})
  expect(compileNativeSheetPagePreviewV1(geometries[0]!,objects).pages).toHaveLength(1)
  let result:unknown='not returned'
  expect(()=>{result=compileNativeSheetPrintAreaSetPreviewV1(geometries,objects)}).toThrow('cannot be met')
  expect(result).toBe('not returned')
 })
 it('applies fit separately and refuses a later incompatible title range without returning partial pages',()=>{
  const selected=[{row:1,column:1,end_row:6,end_column:6},{row:9,column:1,end_row:10,end_column:2}]
  const {objects,compile,part}=fixture(selected),geometries=selected.map(compile)
  Object.assign(objects.page_settings![0]!.settings!,{left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3,fit_to_page:{width:1,height:1}})
  const result=compileNativeSheetPrintAreaSetPreviewV1(geometries,objects)
  expect(result.total_pages).toBe(2);expect(result.areas[0]!.plan.pages[0]!.scale).toBeLessThan(result.areas[1]!.plan.pages[0]!.scale)
  objects.print_titles=[{sheet_id:'7',sheet_part:part,status:'available',rows:{start:1,end:1},warnings:['Source titles']}]
  expect(()=>compileNativeSheetPagePreviewV1(geometries[0]!,objects,undefined,{repeat_print_titles:true})).not.toThrow()
  expect(()=>compileNativeSheetPrintAreaSetPreviewV1(geometries,objects,undefined,{repeat_print_titles:true})).toThrow('must lead')
 })
 it('does not execute accessors or custom iterators on the geometry array',()=>{
  const {objects,compile}=fixture(),geometries=ranges.map(compile)
  let called=false
  const accessor=[...geometries];Object.defineProperty(accessor,'1',{enumerable:true,get(){called=true;return geometries[1]}})
  expect(()=>compileNativeSheetPrintAreaSetPreviewV1(accessor,objects)).toThrow('plain array');expect(called).toBe(false)
  const iterator=[...geometries];Object.defineProperty(iterator,Symbol.iterator,{value(){called=true;return geometries[Symbol.iterator]()}})
  expect(()=>compileNativeSheetPrintAreaSetPreviewV1(iterator,objects)).toThrow('plain array');expect(called).toBe(false)
  const sparse=Array(2);sparse[0]=geometries[0]
  expect(()=>compileNativeSheetPrintAreaSetPreviewV1(sparse,objects)).toThrow('plain array')
 })
 it('enforces 100 pages across otherwise individually valid areas',()=>{
  const selected=[{row:0,column:0,end_row:2999,end_column:1},{row:3000,column:0,end_row:5999,end_column:1}]
  const {objects,compile}=fixture(selected),geometries=selected.map(compile)
  const counts=geometries.map(g=>compileNativeSheetPagePreviewV1(g,objects).pages.length)
  expect(counts.every(n=>n<=100)).toBe(true);expect(counts.reduce((a,b)=>a+b,0)).toBeGreaterThan(100)
  expect(()=>compileNativeSheetPrintAreaSetPreviewV1(geometries,objects)).toThrow('aggregate 100-page')
 })
})
