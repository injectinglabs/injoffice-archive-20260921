import {createRequire} from 'node:module'
import {createHash} from 'node:crypto'
import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {projectNativeWorkbookV2,createNativeMaximumDigitWidthAuthorityV2,compileNativeSheetGeometryV2,compileNativeStoredRowSheetGeometryV1,isCompiledNativeSheetGeometryV2,validateNativeSheetGeometryV2,compileNativeSheetPagePreviewV1,selectNativeSheetPrintTitleViewportV1,type NativeWorkbookV2,type NativeWorkbookObjectsV1,type NativeSheetPreviewPageV1} from './index.js'
const require=createRequire(import.meta.url)
const column=(index:number,width:number)=>({column:index,end_column:index,width,hidden:false,custom_width:true,best_fit:false})
function fixture(change?:(workbook:NativeWorkbookV2)=>void){
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8')) as NativeWorkbookV2
 Object.assign(workbook,{normal_style:{style_xf_id:0,font_id:0,font_name:'DejaVu Sans',font_size_points:11,font_bold:false,font_italic:false,font_record_sha256:`sha256:${'e'.repeat(64)}`}})
 change?.(workbook)
 const model=projectNativeWorkbookV2(workbook),font=new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
 const geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:2,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font))
 const objects:NativeWorkbookObjectsV1={protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:geometry.source_package_sha256,tables:[],charts:[],page_settings:[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',warnings:['Source settings'],settings:{paper:'Letter',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}}]}
 return {geometry,objects,model,font}
}
describe('source-bound selected worksheet page geometry',()=>{
 it('repeats both leading title axes with disjoint regions and complete body coverage',()=>{
  const {model,font,objects}=fixture(),geometry=compileNativeSheetGeometryV2(model,'7',{row:1,column:1,end_row:6,end_column:6},createNativeMaximumDigitWidthAuthorityV2(model,font))
  objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:1,end:1},columns:{start:1,end:1},warnings:['Saved titles']}]
  const baseline=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(baseline.pages.every(p=>p.regions===undefined)).toBe(true)
  Object.assign(objects.page_settings![0]!.settings!,{left_inches:3,right_inches:3,top_inches:5,bottom_inches:5})
  const before=JSON.stringify(objects),p=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})
  expect(p.pages.length).toBeGreaterThan(1)
  const allBody=new Set<string>()
  for(const page of p.pages){
   expect(page.regions!.map(r=>r.kind)).toEqual(['body','repeat-rows','repeat-columns','repeat-corner'])
   const cells=new Set<string>()
   for(const r of page.regions!){
    for(let row=r.rows.start;row<=r.rows.end;row++)for(let col=r.columns.start;col<=r.columns.end;col++){
     const key=`${row}:${col}`;expect(cells.has(key)).toBe(false);cells.add(key)
     if(r.kind==='body'){expect(allBody.has(key)).toBe(false);allBody.add(key)}
    }
    const right=(r.source_clip.x_emu+r.source_clip.width_emu)*page.scale+r.translate_x_emu,bottom=(r.source_clip.y_emu+r.source_clip.height_emu)*page.scale+r.translate_y_emu
    expect(right).toBeLessThanOrEqual(page.content_clip.x_emu+page.content_clip.width_emu)
    expect(bottom).toBeLessThanOrEqual(page.content_clip.y_emu+page.content_clip.height_emu)
   }
  }
  expect(allBody.size).toBe(25);expect(JSON.stringify(objects)).toBe(before)
  Object.assign(objects.page_settings![0]!.settings!,{fit_to_page:{width:1,height:1}})
  const fit=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})
  expect(fit.pages).toHaveLength(1);expect(fit.pages[0]!.scale).toBeLessThan(1)
  expect(fit.pages[0]!.regions).toHaveLength(4)
 })
 it('rejects missing, stale, outside-geometry, all-title and hidden-only repetition',()=>{
  const {geometry,objects,model,font}=fixture(),options={repeat_print_titles:true} as const
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,options)).toThrow('titles unavailable')
  objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:4,end:4},warnings:['Source']}]
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,options)).toThrow('fully contained')
  objects.print_titles[0]!.rows={start:0,end:2}
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,options)).toThrow('leave body')
  objects.print_titles[0]!.rows={start:0,end:0};objects.print_titles[0]!.sheet_part='wrong.xml'
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,options)).toThrow('worksheet part')
  objects.print_titles[0]!.sheet_part='Worksheets/Sheet1.xml'
  objects.row_geometry=[{sheet_part:'Worksheets/Sheet1.xml',rows:Array.from({length:32},(_,row)=>({row,height_points:14.4,hidden:row===0})),warnings:['Stored']}]
  const hidden=compileNativeStoredRowSheetGeometryV1(model,'7',{row:0,column:0,end_row:2,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font),objects)
  expect(()=>compileNativeSheetPagePreviewV1(hidden,objects,undefined,options)).toThrow('visible')
  for(const bad of [null,{},false,{repeat_print_titles:false},{repeat_print_titles:true,extra:1}])expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,bad as any)).toThrow()
 })
 it('selects source-qualified heading geometry outside the body without printing the gap',()=>{
  const {model,font,objects}=fixture(),body={row:5,column:4,end_row:8,end_column:7}
  objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:1,end:1},columns:{start:1,end:1},warnings:['Source']}]
  const view=selectNativeSheetPrintTitleViewportV1(model,'7',body,objects)
  expect(view).toEqual({row:1,column:1,end_row:8,end_column:7});expect(Object.isFrozen(view)).toBe(true)
  const geometry=compileNativeSheetGeometryV2(model,'7',view,createNativeMaximumDigitWidthAuthorityV2(model,font))
  const plan=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true,body_viewport:body})
  expect(plan.pages).toHaveLength(1)
  const [b,r,c,corner]=plan.pages[0]!.regions!
  expect(b!.rows).toEqual({start:5,end:8});expect(b!.columns).toEqual({start:4,end:7})
  expect(r!.rows).toEqual({start:1,end:1});expect(r!.columns).toEqual(b!.columns)
  expect(c!.columns).toEqual({start:1,end:1});expect(c!.rows).toEqual(b!.rows)
  expect(corner!.rows).toEqual(r!.rows);expect(corner!.columns).toEqual(c!.columns)
  for(const region of plan.pages[0]!.regions!) {
   expect(region.rows.start).not.toBe(2);expect(region.columns.start).not.toBe(2)
   expect(region.source_clip.x_emu*plan.pages[0]!.scale+region.translate_x_emu).toBeGreaterThanOrEqual(plan.pages[0]!.content_clip.x_emu)
  }
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true,body_viewport:{...body,row:0}})).toThrow('contained')
  expect(()=>selectNativeSheetPrintTitleViewportV1({...model},'7',body,objects)).toThrow('projected')
  expect(()=>selectNativeSheetPrintTitleViewportV1(model,'7',body,{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
  expect(()=>selectNativeSheetPrintTitleViewportV1(model,'8',body,objects)).toThrow('worksheet part')
  expect(()=>selectNativeSheetPrintTitleViewportV1(model,'7',{...body,end_row:5000},objects)).toThrow()
  let called=false;const accessor={...body};Object.defineProperty(accessor,'row',{get(){called=true;return 5}})
  expect(()=>selectNativeSheetPrintTitleViewportV1(model,'7',accessor,objects)).toThrow();expect(called).toBe(false)
 })
 it('keeps nonleading title bands separate from body segments on both axes',()=>{
  const {geometry,objects}=fixture()
  objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:1,end:1},columns:{start:1,end:1},warnings:['Source']}]
  const plan=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})
  expect(plan.pages).toHaveLength(4)
  const body=plan.pages.map(p=>p.regions![0]!)
  expect(body.map(r=>[r.rows.start,r.columns.start])).toEqual([[0,0],[2,0],[0,2],[2,2]])
  expect(body.every(r=>r.rows.start===r.rows.end&&r.columns.start===r.columns.end)).toBe(true)
  for(const p of plan.pages)expect(p.regions!.filter(r=>r.kind==='repeat-corner')).toHaveLength(1)
  objects.page_settings![0]!.settings!.fit_to_page={width:1,height:1}
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})).toThrow('cannot be met')
 })
 it('supports headings after the body and rejects malformed explicit body options',()=>{
  const {model,font,objects}=fixture(),body={row:0,column:0,end_row:1,end_column:1}
  objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:4,end:4},warnings:['Source']}]
  const view=selectNativeSheetPrintTitleViewportV1(model,'7',body,objects)
  const geometry=compileNativeSheetGeometryV2(model,'7',view,createNativeMaximumDigitWidthAuthorityV2(model,font))
  const p=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true,body_viewport:body})
  expect(p.pages[0]!.regions!.map(r=>r.rows)).toEqual([{start:0,end:1},{start:4,end:4}])
  for(const v of [undefined,null,{...body,extra:1},{...body,row:-0},{...body,end_row:0.5}])expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true,body_viewport:v} as any)).toThrow()
 })
 it('ignores omitted-gap merges and refuses merges crossing gap-to-body or gap-to-heading boundaries',()=>{
  for(const spec of [['E3:E4',2,3,false],['E5:E6',4,5,true],['E2:E3',1,2,true]] as const){
   const {model,font,objects}=fixture(workbook=>{
    Object.assign(workbook.sheets[0]!,{cells:workbook.sheets[0]!.cells.filter(c=>c.ref==='F1'),merged_ranges:[{ref:spec[0],row:spec[1],column:4,end_row:spec[2],end_column:4,editable:false}]})
    const location=['MERGED_CELLS','merges','sheet:7','Worksheets/Sheet1.xml','',''].join('\0')
    Object.assign(workbook,{unsupported:[...workbook.unsupported,{id:`unsupported:${createHash('sha256').update(location).digest('hex')}`,code:'MERGED_CELLS',capability:'merges',scope_id:'sheet:7',part_name:'Worksheets/Sheet1.xml',preservation:'preserve-exact',message:'Merged source geometry'}]})
   })
   objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:1,end:1},warnings:['Source']}]
   const body={row:5,column:4,end_row:8,end_column:7},view=selectNativeSheetPrintTitleViewportV1(model,'7',body,objects)
   const geometry=compileNativeSheetGeometryV2(model,'7',view,createNativeMaximumDigitWidthAuthorityV2(model,font))
   const run=()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true,body_viewport:body})
   if(spec[3])expect(run).toThrow('region boundary');else expect(run().pages).toHaveLength(1)
  }
 })
 it('supports single-axis repetition and refuses oversized title reservations',()=>{
  for(const axis of ['rows','columns'] as const){
   const {geometry,objects}=fixture();objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',[axis]:{start:0,end:0},warnings:['Source']}]
   const p=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})
   expect(p.pages[0]!.regions!.map(r=>r.kind)).toEqual(['body',axis==='rows'?'repeat-rows':'repeat-columns'])
   Object.assign(objects.page_settings![0]!.settings!,{left_inches:4.2,right_inches:4.2,top_inches:5.49,bottom_inches:5.49})
   expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})).toThrow('exceeds one page')
  }
 })
 it('keeps heading merges intact, splits one across column pages and refuses merges crossing heading/body partitions',()=>{
  const merged=(ref:string,row:number,column:number,end_row:number,end_column:number)=>fixture(workbook=>{
   Object.assign(workbook.sheets[0]!,{cells:workbook.sheets[0]!.cells.filter(c=>c.ref==='F1'),merged_ranges:[{ref,row,column,end_row,end_column,editable:false}]})
   const location=['MERGED_CELLS','merges','sheet:7','Worksheets/Sheet1.xml','',''].join('\0')
   Object.assign(workbook,{unsupported:[...workbook.unsupported,{id:`unsupported:${createHash('sha256').update(location).digest('hex')}`,code:'MERGED_CELLS',capability:'merges',scope_id:'sheet:7',part_name:'Worksheets/Sheet1.xml',preservation:'preserve-exact',message:'Merged source geometry'}]})
  })
  for(const spec of [['A1:B1',0,0,0,1],['A1:A2',0,0,1,0]] as const){
   const {geometry,objects}=merged(spec[0],spec[1],spec[2],spec[3],spec[4])
   objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:0,end:0},warnings:['Saved']}]
   if(spec[0]==='A1:A2')expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})).toThrow('region boundary')
   else {
    const pages=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})
    expect(pages.pages[0]!.regions!.find(r=>r.kind==='repeat-rows')!.columns).toEqual({start:0,end:2})
    Object.assign(objects.page_settings![0]!.settings!,{left_inches:3.5,right_inches:3.5})
    const wide=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})
    expect(wide.pages.length).toBeGreaterThan(1)
    expect(wide.warnings.join(' ')).toContain('Merged cells split by a page boundary: A1:B1 (1 of 1 painted merged ranges)')
    Object.assign(objects.page_settings![0]!.settings!,{fit_to_page:{width:3,height:3}})
    expect(compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true}).pages[0]!.scale).toBeLessThan(1)
   }
  }
  const {model,font,objects}=merged('A1:A2',0,0,1,0)
  objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:2,end:2},warnings:['Saved']}]
  objects.row_geometry=[{sheet_part:'Worksheets/Sheet1.xml',rows:Array.from({length:32},(_,row)=>({row,height_points:14.4,hidden:row===0})),warnings:['Stored']}]
  const hidden=compileNativeStoredRowSheetGeometryV1(model,'7',{row:0,column:0,end_row:2,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font),objects)
  expect(()=>compileNativeSheetPagePreviewV1(hidden,objects,undefined,{repeat_print_titles:true})).not.toThrow()
 })
 it('retains the page budget and fit lower bound with repeated headings',()=>{
  const {model,font,objects}=fixture(),geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:3999,end_column:24},createNativeMaximumDigitWidthAuthorityV2(model,font))
  objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',rows:{start:0,end:0},columns:{start:0,end:0},warnings:['Saved']}]
  Object.assign(objects.page_settings![0]!.settings!,{fit_to_page:{width:100,height:0}})
  const pages=compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true}).pages
  expect(pages.length).toBeLessThanOrEqual(100);expect(Math.max(...pages.map(p=>p.rows.end))).toBe(3999)
  Object.assign(objects.page_settings![0]!.settings!,{left_inches:4.249,right_inches:4.249,fit_to_page:{width:1,height:0}})
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,undefined,{repeat_print_titles:true})).toThrow('cannot be met')
 })
 // Excel 16.112.4's own export of the local hard-v2 workbook
 // cell-anchored-hidden-shapes prints four pages for a sheet whose 19 full-width
 // A:H merges all straddle the column break: each merge is painted on both
 // column bands, the second repeating its fill from that page's left edge and
 // re-placing its text at the merge origin off that edge. So a split merge is
 // pages plus a disclosure, not a refusal.
 it('splits a merged cell across a page boundary, names it, and keeps one source rectangle',()=>{
  const {model,font,objects}=fixture(workbook=>{
   Object.assign(workbook.sheets[0]!,{cells:workbook.sheets[0]!.cells.filter(c=>c.ref==='F1'),merged_ranges:[{ref:'A1:C1',row:0,column:0,end_row:0,end_column:2,editable:false}]})
   const location=['MERGED_CELLS','merges','sheet:7','Worksheets/Sheet1.xml','',''].join('\0')
   Object.assign(workbook,{unsupported:[...workbook.unsupported,{id:`unsupported:${createHash('sha256').update(location).digest('hex')}`,code:'MERGED_CELLS',capability:'merges',scope_id:'sheet:7',part_name:'Worksheets/Sheet1.xml',preservation:'preserve-exact',message:'Merged source geometry'}]})
  })
  const geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:2,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font))
  const whole=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(whole.pages).toHaveLength(1)
  expect(whole.warnings.some(w=>w.includes('split by a page boundary'))).toBe(false)
  Object.assign(objects.page_settings![0]!.settings!,{left_inches:3.5,right_inches:3.5})
  const before=JSON.stringify(objects),p=compileNativeSheetPagePreviewV1(geometry,objects)
  const m=geometry.merged_ranges[0]!.rect
  const overlaps=p.pages.filter(({source_clip:c})=>m.x_emu<c.x_emu+c.width_emu&&m.x_emu+m.width_emu>c.x_emu&&m.y_emu<c.y_emu+c.height_emu&&m.y_emu+m.height_emu>c.y_emu)
  expect(overlaps.length).toBeGreaterThan(1)
  expect(p.pages.some(({source_clip:c})=>m.x_emu>=c.x_emu&&m.x_emu+m.width_emu<=c.x_emu+c.width_emu)).toBe(false)
  // One source rectangle, still exactly its bands: the merge is not moved,
  // duplicated whole or resized to fit a page, the pages stay whole-band, and
  // the policy is unchanged because band layout is unchanged.
  expect(geometry.merged_ranges).toHaveLength(1)
  expect(m.x_emu).toBe(geometry.columns[0]!.x_emu)
  expect(m.width_emu).toBe(geometry.columns[2]!.x_emu+geometry.columns[2]!.width_emu-geometry.columns[0]!.x_emu)
  expect(p.policy).toBe('whole-bands-down-then-over-v1')
  expect(p.warnings.join(' ')).toContain('Merged cells split by a page boundary: A1:C1 (1 of 1 painted merged ranges)')
  expect(JSON.stringify(objects)).toBe(before)
 })
 it('searches lower scales when a merged cell would cross an otherwise valid fit boundary',()=>{
  const {geometry,objects}=fixture(workbook=>{
   Object.assign(workbook.sheets[0]!,{merged_ranges:[{ref:'B2:C3',row:1,column:1,end_row:2,end_column:2,editable:false}]})
   const location=['MERGED_CELLS','merges','sheet:7','Worksheets/Sheet1.xml','',''].join('\0')
   Object.assign(workbook,{unsupported:[...workbook.unsupported,{id:`unsupported:${createHash('sha256').update(location).digest('hex')}`,code:'MERGED_CELLS',capability:'merges',scope_id:'sheet:7',part_name:'Worksheets/Sheet1.xml',preservation:'preserve-exact',message:'Merged source geometry'}]})
  })
  Object.assign(objects.page_settings![0]!.settings!,{left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3,fit_to_page:{width:3,height:3}})
  const p=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(p.pages[0]!.scale).toBeLessThan(1)
  const m=geometry.merged_ranges[0]!.rect
  expect(p.pages.some(({source_clip:c})=>m.x_emu>=c.x_emu&&m.y_emu>=c.y_emu&&m.x_emu+m.width_emu<=c.x_emu+c.width_emu&&m.y_emu+m.height_emu<=c.y_emu+c.height_emu)).toBe(true)
 })
 it('bounds the largest supported viewport search and enforces the 100-page budget',()=>{
  const {model,font,objects}=fixture(),geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:3999,end_column:24},createNativeMaximumDigitWidthAuthorityV2(model,font))
  Object.assign(objects.page_settings![0]!.settings!,{fit_to_page:{width:100,height:0}})
  const p=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(p.pages.length).toBeLessThanOrEqual(100);expect(p.pages[0]!.scale).toBeLessThan(1)
  expect(Math.max(...p.pages.map(p=>p.rows.end))).toBe(3999)
  expect(Math.max(...p.pages.map(p=>p.columns.end))).toBe(24)
 })
 it('fits actual whole bands using the largest shrink percentage and retains source metadata',()=>{
  const {geometry,objects}=fixture(),settings=objects.page_settings![0]!.settings!
  Object.assign(settings,{scale:400,left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3,fit_to_page:{width:1,height:1}})
  const before=JSON.stringify(objects),p=compileNativeSheetPagePreviewV1(geometry,objects),page=p.pages[0]!
  expect(p.pages).toHaveLength(1);expect(page.scale).toBeLessThan(1);expect(page.scale).toBeGreaterThanOrEqual(.1)
  expect(page.source_clip).toEqual(geometry.bounds);expect(p.settings.scale).toBe(400)
  expect(page.source_clip.width_emu*page.scale).toBeLessThanOrEqual(page.content_clip.width_emu)
  expect(page.source_clip.height_emu*page.scale).toBeLessThanOrEqual(page.content_clip.height_emu)
  expect(geometry.bounds.width_emu*(page.scale+.01)>page.content_clip.width_emu||geometry.bounds.height_emu*(page.scale+.01)>page.content_clip.height_emu).toBe(true)
  expect(p.warnings.join(' ')).toContain("not Excel's fit algorithm");expect(JSON.stringify(objects)).toBe(before)
 })
 it('treats zero as unconstrained, never enlarges, and accepts explicit host fit independently',()=>{
  const {geometry,objects}=fixture(),settings=objects.page_settings![0]!.settings!
  Object.assign(settings,{fit_to_page:{width:1,height:0},scale:10})
  expect(compileNativeSheetPagePreviewV1(geometry,objects).pages[0]!.scale).toBe(1)
  const p=compileNativeSheetPagePreviewV1(geometry,objects,{...settings,kind:'explicit-host-page-policy-v1',left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3,fit_to_page:{width:0,height:1}})
  expect(new Set(p.pages.map(p=>p.rows.start)).size).toBe(1)
  expect(new Set(p.pages.map(p=>p.columns.start)).size).toBeGreaterThan(1)
  expect(p.settings_origin).toBe('explicit-host')
 })
 it('refuses malformed fit choices, accessors and source mismatches without invoking getters',()=>{
  const invalid=[{},null,undefined,{width:0,height:0},{width:-0,height:1},{width:-1,height:1},{width:101,height:1},{width:1.5,height:1},{width:'1',height:1},{width:1,height:Infinity},{width:1,height:1,extra:1},Object.assign(Object.create({}),{width:1,height:1})]
  let invoked=false
  invalid.push(Object.defineProperty({height:1},'width',{enumerable:true,get(){invoked=true;return 1}}))
  for(const fit of invalid){
   const {geometry,objects}=fixture();Object.assign(objects.page_settings![0]!.settings!,{fit_to_page:fit})
   expect(()=>compileNativeSheetPagePreviewV1(geometry,objects)).toThrow()
   expect(()=>compileNativeSheetPagePreviewV1(geometry,fixture().objects,{...objects.page_settings![0]!.settings!,kind:'explicit-host-page-policy-v1'})).toThrow()
  }
  expect(invoked).toBe(false)
  const {geometry,objects}=fixture();objects.page_settings![0]!.sheet_part='wrong.xml'
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects)).toThrow('worksheet part')
 })
 it('refuses targets below the minimum scale rather than silently clipping or missing targets',()=>{
  const {geometry,objects}=fixture();Object.assign(objects.page_settings![0]!.settings!,{left_inches:4.249,right_inches:4.249,fit_to_page:{width:1,height:0}})
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects)).toThrow('cannot be met')
 })
 it('fits non-A1 geometry and transforms viewport-local drawing coordinates uniformly',()=>{
  const {model,font,objects}=fixture(),geometry=compileNativeSheetGeometryV2(model,'7',{row:1,column:1,end_row:3,end_column:3},createNativeMaximumDigitWidthAuthorityV2(model,font))
  Object.assign(objects.page_settings![0]!.settings!,{left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3,fit_to_page:{width:1,height:1}})
  const page=compileNativeSheetPagePreviewV1(geometry,objects).pages[0]!
  expect(page.rows).toEqual({start:1,end:3});expect(page.columns).toEqual({start:1,end:3})
  expect(page.source_clip.x_emu*page.scale+page.translate_x_emu).toBe(page.content_clip.x_emu)
  const chart={x:geometry.bounds.width_emu/4,y:geometry.bounds.height_emu/4,width:geometry.bounds.width_emu/2,height:geometry.bounds.height_emu/2}
  const painted={x:chart.x*page.scale+page.translate_x_emu,y:chart.y*page.scale+page.translate_y_emu,width:chart.width*page.scale,height:chart.height*page.scale}
  expect(painted.x+painted.width).toBeLessThanOrEqual(page.content_clip.x_emu+page.content_clip.width_emu)
  expect(painted.y+painted.height).toBeLessThanOrEqual(page.content_clip.y_emu+page.content_clip.height_emu)
 })
 it('honors source and explicit host page order without changing page geometry or source settings',()=>{
  const {geometry,objects}=fixture()
  Object.assign(objects.page_settings![0]!.settings!,{left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3,page_order:'overThenDown'})
  const before=JSON.stringify(objects),source=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(source.policy).toBe('whole-bands-over-then-down-v1');expect(source.settings_origin).toBe('source')
  expect(source.pages.map(p=>[p.rows.start,p.columns.start])).toEqual([[0,0],[0,1],[0,2],[1,0],[1,1],[1,2],[2,0],[2,1],[2,2]])
  const host=compileNativeSheetPagePreviewV1(geometry,objects,{...objects.page_settings![0]!.settings!,kind:'explicit-host-page-policy-v1',page_order:'downThenOver'})
  expect(host.policy).toBe('whole-bands-down-then-over-v1');expect(host.settings_origin).toBe('explicit-host')
  for(const p of source.pages){const same=host.pages.find(q=>q.rows.start===p.rows.start&&q.columns.start===p.columns.start)!;expect({...p,number:0}).toEqual({...same,number:0})}
  expect(JSON.stringify(objects)).toBe(before)
 })
 it('rejects unknown, empty and non-string ordering rather than silently using defaults',()=>{
  for(const order of ['diagonal','',null,0,undefined]){
   const {geometry,objects}=fixture()
   Object.assign(objects.page_settings![0]!.settings!,{page_order:order})
   expect(()=>compileNativeSheetPagePreviewV1(geometry,objects)).toThrow()
  }
 })
 it('rejects accessor and symbol host policies without invoking getters',()=>{
  const {geometry,objects}=fixture();let invoked=false
  const host={kind:'explicit-host-page-policy-v1',paper:'A4',orientation:'portrait',scale:100,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1}
  Object.defineProperty(host,'scale',{enumerable:true,get(){invoked=true;return 100}})
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,host as any)).toThrow();expect(invoked).toBe(false)
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects,{...objects,[Symbol('foreign')]:1} as any)).toThrow()
 })
 it('paginates whole bands down then over without dropping rows or columns',()=>{
  const {geometry,objects}=fixture()
  const p=compileNativeSheetPagePreviewV1(geometry,objects,{kind:'explicit-host-page-policy-v1',paper:'Letter',orientation:'portrait',scale:100,left_inches:3.5,right_inches:3.5,top_inches:5.3,bottom_inches:5.3})
  expect(p.pages).toHaveLength(9)
  expect(p.pages.map(p=>[p.rows.start,p.columns.start])).toEqual([[0,0],[1,0],[2,0],[0,1],[1,1],[2,1],[0,2],[1,2],[2,2]])
  for(const page of p.pages){expect(page.source_clip.x_emu*page.scale+page.translate_x_emu).toBe(page.content_clip.x_emu);expect(page.source_clip.y_emu*page.scale+page.translate_y_emu).toBe(page.content_clip.y_emu)}
 })
 it('brands stored-row approximation separately, preserves stored sizes and refuses stale evidence',()=>{
  const {model,font,objects}=fixture(),metric=createNativeMaximumDigitWidthAuthorityV2(model,font),viewport={row:0,column:0,end_row:2,end_column:2}
  objects.row_geometry=[{sheet_part:'Worksheets/Sheet1.xml',rows:Array.from({length:32},(_,row)=>({row,height_points:14.4,hidden:row===1})),warnings:['Stored only']}]
  const g=compileNativeStoredRowSheetGeometryV1(model,'7',viewport,metric,objects)
  expect(g.rows[0]!.height_emu).toBe(182880);expect(g.rows[1]!.height_emu).toBe(0)
  expect(g.approximation.policy).toBe('source-stored-rows-v1');expect(isCompiledNativeSheetGeometryV2(g)).toBe(false)
  objects.row_geometry[0]!.root_policy='x14ac-descent-only-v1'
  const rootQualified=compileNativeStoredRowSheetGeometryV1(model,'7',viewport,metric,objects)
  expect(rootQualified.approximation.source.root_policy).toBe('x14ac-descent-only-v1')
  expect(JSON.stringify(rootQualified)).not.toBe(JSON.stringify(g))
  expect(()=>compileNativeStoredRowSheetGeometryV1(model,'7',viewport,metric,{...objects,row_geometry:[{...objects.row_geometry![0]!,root_policy:'ignore-anything'}]} as any)).toThrow()
  expect(()=>validateNativeSheetGeometryV2(g)).toThrow()
  expect(compileNativeSheetPagePreviewV1(g,objects).warnings.join(' ')).toContain('baselines are not qualified')
  expect(()=>compileNativeStoredRowSheetGeometryV1(model,'7',{...viewport,end_row:32},metric,objects)).toThrow('first 32')
  objects.row_geometry[0]!.sheet_part='wrong.xml'
  expect(()=>compileNativeStoredRowSheetGeometryV1(model,'7',viewport,metric,objects)).toThrow('unavailable')
 })
 it('places stored geometry inside authored margins without changing source',()=>{
  const {geometry,objects}=fixture(),before=JSON.stringify(objects),p=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(p.settings_origin).toBe('source');expect(p.pages).toHaveLength(1)
  expect(p.pages[0]).toMatchObject({width_emu:7772400,height_emu:10058400,scale:1,translate_x_emu:914400,translate_y_emu:914400,source_clip:geometry.bounds})
  expect(JSON.stringify(objects)).toBe(before)
 })
 it('reserves the deeper of the top and header margins and of the bottom and footer margins',()=>{
  // ECMA-376 §18.3.1.62 measures all six margins from the paper edge, so Excel's
  // body runs from max(top, header) to max(bottom, footer). These are the
  // authored margins of hard-v2 cell-anchored-hidden-shapes.xlsx, whose header
  // margin is deeper than its top margin; ignoring them printed a body
  // 326,304 EMU too tall and lost a whole row band.
  const {model,font,objects}=fixture()
  const geometry=compileNativeSheetGeometryV2(model,'7',{row:0,column:0,end_row:33,end_column:2},createNativeMaximumDigitWidthAuthorityV2(model,font))
  const inch=(value:number)=>Math.round(value*914400)
  const settings=objects.page_settings![0]!.settings!
  Object.assign(settings,{paper:'A4',orientation:'portrait',scale:153,left_inches:0.7,right_inches:0.7,top_inches:0.63,bottom_inches:0,header_inches:0.79,footer_inches:0.19685})
  const reserved=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(reserved.pages[0]!.content_clip.y_emu).toBe(inch(0.79))
  expect(reserved.pages[0]!.content_clip.height_emu).toBe(10692000-inch(0.79)-inch(0.19685))
  expect(reserved.pages).toHaveLength(2)
  // A worksheet that states no header or footer band reserves none, so the
  // body keeps the top and bottom margins it authored.
  delete settings.header_inches;delete settings.footer_inches
  const bare=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(bare.pages[0]!.content_clip.y_emu).toBe(inch(0.63))
  expect(bare.pages[0]!.content_clip.height_emu).toBe(10692000-inch(0.63))
  expect(bare.pages).toHaveLength(1)
  // A header shallower than the top margin changes nothing: the body edge is
  // the deeper of the two, never their sum.
  Object.assign(settings,{header_inches:0.3,footer_inches:0})
  const shallow=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(shallow.pages[0]!.content_clip).toEqual(bare.pages[0]!.content_clip)
  // An explicit host policy may state the bands too, and is still exact data.
  const host=compileNativeSheetPagePreviewV1(geometry,objects,{kind:'explicit-host-page-policy-v1',paper:'A4',orientation:'portrait',scale:153,left_inches:0.7,right_inches:0.7,top_inches:0.63,bottom_inches:0,header_inches:0.79,footer_inches:0.19685})
  expect(host.pages[0]!.content_clip).toEqual(reserved.pages[0]!.content_clip)
  expect(host.settings.header_inches).toBe(0.79)
 })
 it('requires explicit host choice for unavailable settings and labels it',()=>{
  const {geometry,objects}=fixture();objects.page_settings![0]={sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'unavailable',warnings:['Printer settings not modeled']}
  expect(()=>compileNativeSheetPagePreviewV1(geometry,objects)).toThrow('explicit host')
  const p=compileNativeSheetPagePreviewV1(geometry,objects,{kind:'explicit-host-page-policy-v1',paper:'A4',orientation:'landscape',scale:50,left_inches:1,right_inches:1,top_inches:1,bottom_inches:1})
  expect(p.settings_origin).toBe('explicit-host');expect(p.pages[0]).toMatchObject({width_emu:10692000,height_emu:7560000,scale:.5})
  expect(p.warnings.join(' ')).toContain('not authored')
 })
 it('refuses stale and forged geometry',()=>{
  const {geometry,objects}=fixture()
  expect(()=>compileNativeSheetPagePreviewV1({...geometry},objects)).toThrow('compiled')
  expect(()=>compileNativeSheetPagePreviewV1(geometry,{...objects,package_sha256:`sha256:${'b'.repeat(64)}`})).toThrow()
 })
 // Excel breaks pages on whole column boundaries and clips the single column
 // that alone exceeds the printable area rather than dropping the page, so the
 // oversize band keeps a whole-band range and only its painted extent is cut.
 it('clips a single oversized column into its own whole-band page and names the cut',()=>{
  const {geometry,objects}=fixture(workbook=>{Object.assign(workbook.sheets[0]!,{columns:[...workbook.sheets[0]!.columns,column(2,100)]})})
  const cw=7772400-2*914400
  const oversize=geometry.columns.find(c=>c.width_emu>cw)!
  expect(oversize.column).toBe(2)
  const p=compileNativeSheetPagePreviewV1(geometry,objects)
  expect(p.policy).toBe('whole-bands-down-then-over-clipped-oversize-band-v1')
  expect(p.pages).toHaveLength(2)
  const [first,second]=p.pages as [NativeSheetPreviewPageV1,NativeSheetPreviewPageV1]
  expect(first.columns).toEqual({start:0,end:1});expect(second.columns).toEqual({start:2,end:2})
  expect(first.rows).toEqual(second.rows)
  expect(first.source_clip.width_emu).toBe(geometry.columns[0]!.width_emu+geometry.columns[1]!.width_emu)
  expect(second.source_clip).toEqual({x_emu:oversize.x_emu,y_emu:0,width_emu:cw,height_emu:first.source_clip.height_emu})
  // The page body itself is untouched and the clipped band still starts at the left margin.
  expect(second.content_clip).toEqual(first.content_clip)
  expect(second.translate_x_emu+second.source_clip.x_emu*second.scale).toBe(914400)
  expect(p.warnings.join(' ')).toContain(`Column 3 is ${oversize.width_emu} EMU where this page allows ${cw}, so ${oversize.width_emu-cw} EMU are cut from its trailing edge`)
  expect(p.warnings.join(' ')).toContain('not recoverable from this preview')
 })
 it('refuses an oversized repeated print-title band and never reaches the clip from fit-to-page',()=>{
  // A clipped repeated title would be truncated on every page it repeats onto.
  const title=fixture(workbook=>{Object.assign(workbook.sheets[0]!,{columns:[column(0,100),column(1,12.5)]})})
  title.objects.print_titles=[{sheet_id:'7',sheet_part:'Worksheets/Sheet1.xml',status:'available',columns:{start:0,end:0},warnings:['Source']}]
  expect(()=>compileNativeSheetPagePreviewV1(title.geometry,title.objects,undefined,{repeat_print_titles:true})).toThrow('exceeds one page')
  // Fit-to-page only ever selects a shrink that needs no clipping, and refuses when none exists.
  const wide=()=>{
   const f=fixture(workbook=>{Object.assign(workbook.sheets[0]!,{columns:[...workbook.sheets[0]!.columns,column(2,130)]})})
   Object.assign(f.objects.page_settings![0]!.settings!,{left_inches:3.8,right_inches:3.8})
   return f
  }
  const fit=wide();Object.assign(fit.objects.page_settings![0]!.settings!,{fit_to_page:{width:1,height:1}})
  expect(()=>compileNativeSheetPagePreviewV1(fit.geometry,fit.objects)).toThrow('Fit-to-page target cannot be met')
  const authored=wide()
  expect(compileNativeSheetPagePreviewV1(authored.geometry,authored.objects).policy).toBe('whole-bands-down-then-over-clipped-oversize-band-v1')
 })
})
