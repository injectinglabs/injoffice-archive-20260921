import {createHash} from 'node:crypto'
import {createRequire} from 'node:module'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {
  compileNativeSheetGeometryV2,
  compileNativeSheetPrintPagePreviewV1,
  createNativeMaximumDigitWidthAuthorityV2,
  decodeNativeSheetDimensionNeutralityV1,
  NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_CODES,
  NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_POLICY,
  projectNativeWorkbookV2,
  type NativeSheetDimensionNeutralityV1,
  type NativeSheetViewportV2,
  type NativeWorkbookObjectsV1,
  type NativeWorkbookV2,
} from './index.js'

const require=createRequire(import.meta.url)
const area:NativeSheetViewportV2={row:1,column:1,end_row:3,end_column:3}
const PART='Worksheets/Sheet1.xml'
// Every disclosure the bounded dimension projection currently gates on, as the
// extractor writes them for a modern Excel worksheet.
const GATED=[
 {code:'SHEET_VIEW_GEOMETRY',capability:'dimensions',message:'unmodeled sheet-view geometry remains authority-bound to the source package'},
 {code:'SHEET_FORMAT_EXTRAS',capability:'dimensions',message:'outline levels, thick-edge flags, or unmodeled sheet-format attributes remain authority-bound to the source package'},
 {code:'ROW_DIMENSION_EXTRAS',capability:'dimensions',message:'spans, outline, collapsed, custom-format, thick-border, phonetic, or unmodeled row attributes are preserved'},
 {code:'COLUMN_DIMENSION_EXTRAS',capability:'dimensions',message:'outline, collapsed, phonetic, or unmodeled column attributes are preserved'},
 {code:'COLS_ATTRIBUTES',capability:'dimensions',message:'unmodeled column-container attributes are preserved exactly'},
 {code:'WORKSHEET_ATTRIBUTES',capability:'worksheet-features',message:'unmodeled worksheet attributes remain authority-bound to the source package'},
 {code:'FOREIGN_WORKSHEET_MARKUP',capability:'extensions',message:'foreign worksheet markup remains source-authoritative'},
] as const

function unsupportedID(code:string,capability:string):string{
 const location=[code,capability,'sheet:7',PART,'',''].join('\0')
 return `unsupported:${createHash('sha256').update(location).digest('hex')}`
}
function fixture(disclosures:readonly{code:string;capability:string;message:string}[]){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 Object.assign(workbook,{unsupported:[
  ...workbook.unsupported,
  ...disclosures.map(item=>({
   id:unsupportedID(item.code,item.capability),
   code:item.code,capability:item.capability,scope_id:'sheet:7',part_name:PART,
   preservation:'preserve-exact' as const,message:item.message,
  })),
 ]})
 // WORKSHEET_ATTRIBUTES is a mutation refusal in the contract; the read-only
 // dimension projection is the only thing this evidence unblocks.
 if(disclosures.some(item=>item.code==='WORKSHEET_ATTRIBUTES')){
  Object.assign(workbook.sheets[0]!,{editable:false,refusal_code:'UNSAFE_WORKSHEET_ATTRIBUTES'})
 }
 const model=projectNativeWorkbookV2(workbook)
 const font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
 return {model,authority:createNativeMaximumDigitWidthAuthorityV2(model,font)}
}
function objects(model:ReturnType<typeof projectNativeWorkbookV2>,neutrality?:NativeSheetDimensionNeutralityV1[]):NativeWorkbookObjectsV1{
 return {
  protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:model.source.package_sha256,tables:[],charts:[],
  print_area_sets:[{sheet_id:'7',sheet_part:PART,status:'available',areas:[{...area}],warnings:['Saved source rectangle']}],
  page_settings:[{sheet_id:'7',sheet_part:PART,status:'available',warnings:['Source settings'],settings:{paper:'Letter',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}}],
  ...(neutrality?{dimension_neutrality:neutrality}:{}),
 }
}
function evidence(overrides:Partial<NativeSheetDimensionNeutralityV1>={}):NativeSheetDimensionNeutralityV1{
 return {
  sheet_part:PART,
  policy:NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_POLICY,
  codes:[...NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_CODES],
  warnings:['Non-dimensional worksheet markup; the bounded dimension projection is unchanged.'],
  ...overrides,
 }
}

describe('non-dimensional worksheet markup evidence',()=>{
 it('refuses the bounded dimension projection without evidence',()=>{
  for(const disclosure of GATED){
   const {model,authority}=fixture([disclosure])
   expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority)).toThrow(`source dimension semantics ${disclosure.code} are not projected exactly`)
  }
 })

 it('compiles the same geometry once every gated disclosure is qualified',()=>{
  const clean=fixture([])
  const expected=compileNativeSheetGeometryV2(clean.model,'7',area,clean.authority)
  const {model,authority}=fixture(GATED)
  expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority)).toThrow('are not projected exactly')
  const compiled=compileNativeSheetGeometryV2(model,'7',area,authority,objects(model,[evidence()]))
  expect(compiled.rows).toEqual(expected.rows)
  expect(compiled.columns).toEqual(expected.columns)
  expect(compiled.bounds).toEqual(expected.bounds)
  expect(compiled.merged_ranges).toEqual(expected.merged_ranges)
  expect(compiled.geometry_sha256).toBe(expected.geometry_sha256)
 })

 it('produces the source print-page rectangles the refusal was hiding',()=>{
  const {model,authority}=fixture(GATED)
  const evidenced=objects(model,[evidence()])
  const geometry=compileNativeSheetGeometryV2(model,'7',area,authority,evidenced)
  const preview=compileNativeSheetPrintPagePreviewV1([geometry],evidenced)
  expect(preview.status).toBe('available')
  if(preview.status!=='available')return
  expect(preview.pages).toHaveLength(1)
  expect(preview.pages[0]!.width_css_px).toBe(816)
  expect(preview.pages[0]!.height_css_px).toBe(1056)
 })

 it('waives only the codes the evidence lists',()=>{
  const {model,authority}=fixture(GATED)
  const partial=objects(model,[evidence({codes:['SHEET_VIEW_GEOMETRY','SHEET_FORMAT_EXTRAS']})])
  expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority,partial)).toThrow('ROW_DIMENSION_EXTRAS are not projected exactly')
 })

 // FOREIGN_WORKSHEET_MARKUP joined the policy's list for the one shape the Go
 // tier qualifies: a worksheet whose every foreign child is a
 // markup-compatibility block holding nothing but anchored form controls.
 // Evidence that does not list it leaves that refusal standing, and evidence
 // naming a code the policy never lists is not evidence at all.
 it('waives foreign worksheet markup only when the evidence lists it',()=>{
  const {model,authority}=fixture([{code:'FOREIGN_WORKSHEET_MARKUP',capability:'extensions',message:'foreign worksheet markup remains source-authoritative'}])
  const silent=objects(model,[evidence({codes:['SHEET_VIEW_GEOMETRY']})])
  expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority,silent))
   .toThrow('FOREIGN_WORKSHEET_MARKUP are not projected exactly')
  expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority,objects(model,[evidence()]))).not.toThrow()
 })

 it('never accepts evidence for a code outside the bounded projection vocabulary',()=>{
  for(const code of ['DRAWING_REFERENCE','MERGED_CELLS','WORKSHEET_EXTENSIONS']){
   expect(()=>decodeNativeSheetDimensionNeutralityV1([evidence({codes:[code] as never})])).toThrow()
  }
 })

 it('refuses evidence that does not join this worksheet part',()=>{
  const {model,authority}=fixture(GATED)
  const other=objects(model,[evidence({sheet_part:'Worksheets/Sheet2.xml'})])
  expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority,other)).toThrow('are not projected exactly')
 })

 it('refuses evidence bound to another source package',()=>{
  const {model,authority}=fixture(GATED)
  const foreign={...objects(model,[evidence()]),package_sha256:`sha256:${'b'.repeat(64)}`}
  expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority,foreign)).toThrow(TypeError)
 })

 it('rejects malformed, duplicated or off-policy evidence',()=>{
  const cases:unknown[]=[
   [{...evidence(),policy:'anything-goes'}],
   [{...evidence(),codes:['MERGED_CELLS']}],
   [{...evidence(),codes:['SHEET_VIEW_GEOMETRY','SHEET_VIEW_GEOMETRY']}],
   [{...evidence(),warnings:[]}],
   [{...evidence(),sheet_part:''}],
   [{...evidence(),sheet_part:'../escape.xml'}],
   [{...evidence(),extra:1}],
   [evidence(),evidence()],
   'not-an-array',
   [null],
  ]
  for(const value of cases)expect(()=>decodeNativeSheetDimensionNeutralityV1(value)).toThrow(TypeError)
 })

 it('accepts one entry per worksheet part',()=>{
  const {model,authority}=fixture(GATED)
  const duplicated=objects(model,[evidence(),evidence({sheet_part:'Worksheets/Sheet2.xml'})])
  expect(()=>compileNativeSheetGeometryV2(model,'7',area,authority,duplicated)).not.toThrow()
 })
})
