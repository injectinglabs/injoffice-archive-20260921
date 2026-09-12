import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {createNativeDocxPartialContentPreviewV1 as project,DOCX_PARTIAL_CONTENT_POLICY} from './nativePartialContentV1.js'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
const original=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1
const options={policy:DOCX_PARTIAL_CONTENT_POLICY,read_only:true as const}
function fixture(){
 const document=structuredClone(original);document.unsupported=[]
 const resolved:NativeDocxResolvedLayoutInputV1={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
 for(const block of document.body.blocks){
  const paragraphs=block.paragraph?[block.paragraph]:block.table!.rows.flatMap(row=>row.cells.flatMap(cell=>cell.paragraphs))
  for(const p of paragraphs){resolved.paragraphs.push({paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}});for(const r of p.runs)resolved.runs.push({run_id:r.id,paragraph_id:p.id,applied_paragraph_styles:[],applied_character_styles:[],properties:{}})}
 }
 return {document,resolved}
}
describe('read-only native partial source content',()=>{
 it('retains exact nontext metadata without allowing similarly named or misplaced diagnostics',()=>{
  const {document,resolved}=fixture();resolved.source_parts.styles_part='word/styles.xml'
  const diagnostic={code:'LATENT_STYLE_BEHAVIOR_PRESERVED',scope_id:document.document_id,part_name:'word/styles.xml',path:'/w:styles[1]/w:latentStyles[1]',severity:'unsupported' as const,preservation:'preserve-verbatim' as const,message:'Preserved metadata'}
  resolved.diagnostics=[diagnostic]
  const out=project(document,options,resolved)
  expect(out.coverage.projected_text_runs).toBeGreaterThan(0)
  expect(out.source_diagnostics.resolved).toEqual([diagnostic])
  expect(out.retained_nontext_diagnostic_ids).toEqual(['resolved:0:LATENT_STYLE_BEHAVIOR_PRESERVED'])
  for(const change of [{scope_id:'unknown'},{part_name:'word/other.xml'},{path:'/w:styles[1]/w:style[1]'},{code:'UNKNOWN_STYLE_METADATA'}]){
   resolved.diagnostics=[{...diagnostic,...change}]
   let recovered=0;try{recovered=project(document,options,resolved).coverage.projected_text_runs}catch(error){expect(error).toBeInstanceOf(TypeError)}
   expect(recovered).toBe(0)
  }
 })
 it('qualifies only the joined automatic-border style diagnostic, never arbitrary table effects',()=>{
  const {document,resolved}=fixture(),table=document.body.blocks[1]!.table!
  delete table.borders
  const part=document.passthrough_parts[0]!;resolved.source_parts.styles_part=part.part_name
  const path='/w:styles[1]/w:style[1]/w:tblPr[1]'
  const diagnostic={code:'TABLE_STYLE_EFFECTS_PRESERVED',scope_id:table.id,part_name:part.part_name,path,severity:'unsupported' as const,preservation:'preserve-verbatim' as const,message:'Automatic borders'}
  const evidence={policy:'auto-border-on-qualified-white-v1' as const,read_only:true as const,package_sha256:document.source.package_sha256,page_background:'absent-on-white-preview' as const,background_rgb:'FFFFFF' as const,source_part:part.part_name,source_path:path+'/w:tblBorders[1]',source_sha256:part.sha256,borders:{top:{style:'single' as const,size_eighth_points:4,color_rgb:'000000'}},automatic_edges:['top' as const],cell_ids:table.rows.flatMap(r=>r.cells.map(c=>c.id)),source_diagnostics:[{code:diagnostic.code,scope_id:table.id,part_name:part.part_name,path}]}
  resolved.tables=[{table_id:table.id,automatic_border_preview:evidence}];resolved.diagnostics=[diagnostic]
  expect(JSON.stringify(project(document,options,resolved).blocks)).toContain('Summary')
  const before=structuredClone({document,resolved});project(document,options,resolved);expect({document,resolved}).toEqual(before)
  table.rows[0]!.cells[0]!.grid_span=2
  expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
  table.rows[0]!.cells[0]!.grid_span=1
  for(const change of [{package_sha256:'sha256:'+'0'.repeat(64)},{source_sha256:'sha256:'+'0'.repeat(64)},{cell_ids:['wrong-cell']},{source_path:path+'/w:nested[1]/w:tblBorders[1]'},{source_diagnostics:[{...evidence.source_diagnostics[0]!,path:'/wrong'}]}]){
   resolved.tables[0]!.automatic_border_preview={...evidence,...change}
   expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
  }
  delete resolved.tables[0]!.automatic_border_preview
  expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
 })
 it('recovers ordered unmerged cell text without claiming table geometry',()=>{
  const {document,resolved}=fixture(),out=project(document,options,resolved),table=out.blocks.find(b=>b.kind==='table-source')!
  if(table.kind!=='table-source')throw new Error('Expected cell groups')
  expect(table.cells[0]).toMatchObject({row_ordinal:0,cell_ordinal:0,kind:'cell-source'})
  expect(JSON.stringify(table)).toContain('Summary')
  expect(table).not.toHaveProperty('width');expect(out.omissions.some(o=>o.code==='table')).toBe(true)
  const inventory=project(document,options).blocks.find(b=>b.kind==='table-source')!
  expect(JSON.stringify(inventory)).not.toContain('Summary')
  expect(JSON.stringify(inventory)).toContain('unqualified-text-visibility')
 })
 it('retains drawing omission alongside qualified literal authored alternative text',()=>{
  const {document,resolved}=fixture(),run=document.body.blocks[0]!.paragraph!.runs.find(r=>r.drawing)!
  run.drawing!.alt_text='<script>description</script>'
  const out=project(document,options,resolved),p=out.blocks[0]!
  if(p.kind!=='paragraph')throw new Error('Expected paragraph')
  expect(p.segments).toContainEqual({kind:'alternative-text',source:{scope_id:run.drawing!.id,anchor:run.drawing!.anchor},text:run.drawing!.alt_text,label:'authored-drawing-description'})
  expect(out.omissions.some(o=>o.code==='drawing')).toBe(true)
  expect(JSON.stringify(project(document,options))).not.toContain('<script>description</script>')
  resolved.runs.find(r=>r.run_id===run.id)!.properties.hidden=true
  expect(JSON.stringify(project(document,options,resolved))).not.toContain('<script>description</script>')
 })
 it('keeps continuation and unsupported table-cell content behind explicit omissions',()=>{
  for(const reason of ['merged','cell-diagnostic','row-diagnostic'] as const){
   const {document,resolved}=fixture(),table=document.body.blocks[1]!.table!,cell=table.rows[0]!.cells[0]!
   if(reason==='merged')cell.vertical_merge='continue'
   else {const scope=reason==='cell-diagnostic'?cell:table.rows[0]!;document.unsupported=[{id:'unsafe-cell',code:'UNMODELED',scope_id:scope.id,anchor:scope.anchor,capability:'tables',preservation:'preserve-verbatim',message:'Unknown'}]}
   const out=project(document,options,resolved),group=out.blocks.find(b=>b.kind==='table-source')!
   expect(JSON.stringify(group)).not.toContain('Summary')
   expect(out.omissions.some(o=>o.code===(reason==='merged'?'merged-cell':'unsupported-source'))).toBe(true)
  }
 })
 it.each(['none','restart'] as const)('recovers merged-owner source text with vertical role %s and preserves all authority',vertical=>{
  const {document,resolved}=fixture(),cell=document.body.blocks[1]!.table!.rows[0]!.cells[0]!
  cell.grid_span=2;cell.vertical_merge=vertical
  const before=structuredClone({document,resolved}),out=project(document,options,resolved),table=out.blocks.find(b=>b.kind==='table-source')!
  expect(JSON.stringify(table)).toContain('Summary')
  if(table.kind!=='table-source')throw Error('Expected source table')
  expect(table.cells[0]!.source_merge).toEqual({grid_span:2,vertical_merge:vertical})
  expect(out.omissions.some(o=>o.code==='table')).toBe(true)
  expect({document,resolved}).toEqual(before)
  expect(JSON.stringify(project(document,options).blocks)).not.toContain('Summary')
  const run=cell.paragraphs[0]!.runs[0]!
  resolved.runs.find(r=>r.run_id===run.id)!.properties.hidden=true
  expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
  delete resolved.runs.find(r=>r.run_id===run.id)!.properties.hidden
  run.properties={...run.properties,hidden:true}
  expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
 })
 it('does not broaden nested-table, paragraph, run, table or global diagnostic exceptions for merged owners',()=>{
  for(const kind of ['nested','table','paragraph','run','global'] as const){
   const {document,resolved}=fixture(),table=document.body.blocks[1]!.table!,cell=table.rows[0]!.cells[0]!,p=cell.paragraphs[0]!
   cell.grid_span=2;cell.vertical_merge='restart'
   const owner=kind==='paragraph'?p:kind==='run'?p.runs[0]!:table
   document.unsupported=[{id:'unsafe-owner',code:kind==='nested'?'NESTED_TABLE_OR_CELL_MARKUP':'UNMODELED',scope_id:kind==='global'?document.document_id:owner.id,anchor:owner.anchor,capability:'tables',preservation:'preserve-verbatim',message:'Preserved'}]
   expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
  }
 })
 it('bounds merged-owner table-cell groups and records exact source-cell remainder',()=>{
  const {document}=fixture(),table=document.body.blocks[1]!.table!,original=table.rows[0]!
  original.cells[0]!.grid_span=2;original.cells[0]!.vertical_merge='restart'
  table.rows=Array.from({length:203},(_,i)=>{
   const row=structuredClone(original);row.id=`row:${i}`;row.anchor.path=table.anchor.path+`/w:tr[${i+1}]`
   for(const [j,cell]of row.cells.entries()){cell.id=`cell:${i}:${j}`;cell.anchor.path=row.anchor.path+`/w:tc[${j+1}]`;for(const [k,p]of cell.paragraphs.entries()){p.id=`p:${i}:${j}:${k}`;p.anchor.path=cell.anchor.path+`/w:p[${k+1}]`;for(const [n,r]of p.runs.entries()){r.id=`r:${i}:${j}:${k}:${n}`;r.anchor.path=p.anchor.path+`/w:r[${n+1}]`}}}
   return row
  })
  const out=project(document,options),group=out.blocks.find(b=>b.kind==='table-source')!
  if(group.kind!=='table-source')throw new Error('Expected table groups')
  expect(group.cells).toHaveLength(200);expect(group.source_cell_count).toBe(203)
  expect(group.cells.every(cell=>cell.source_merge?.grid_span===2&&cell.source_merge.vertical_merge==='restart')).toBe(true)
  expect(out.omissions.find(o=>o.code==='cell-limit')).toMatchObject({count:3,source:{scope_id:table.id}})
 })
 it('keeps safe source text alongside explicit drawing/table/story omissions, never page or mutation output',()=>{
  const {document,resolved}=fixture(),before=structuredClone(document),out=project(document,options,resolved)
  expect(out.coverage.projected_text_runs).toBeGreaterThan(0)
  expect(out.omissions.map(o=>o.code)).toEqual(expect.arrayContaining(['drawing','table','nonbody-story']))
  expect(out.pagination).toBe('not-produced');expect(out.read_only).toBe(true)
  expect(out.source.package_sha256).toBe(document.source.package_sha256)
  expect(document).toEqual(before)
  out.blocks.length=0;out.source_diagnostics.document.push({id:'changed',code:'changed',capability:'changed',scope_id:'changed',preservation:'preserve-verbatim',message:'changed'})
  expect(document).toEqual(before)
 })
 it('without layout emits an honest inventory and no potentially inherited hidden text',()=>{
  const {document}=fixture(),out=project(document,options)
  expect(out.coverage.projected_text_runs).toBe(0);expect(out.coverage.layout_present).toBe(false)
  expect(out.omissions.map(o=>o.code)).toEqual(expect.arrayContaining(['unqualified-text-visibility','drawing','table']))
 })
 it('suppresses resolved hidden runs and keeps each omission source-bound',()=>{
  const {document,resolved}=fixture(),p=document.body.blocks[0]!.paragraph!,run=p.runs.find(r=>r.kind==='text')!
  resolved.runs.find(r=>r.run_id===run.id)!.properties.hidden=true
  const out=project(document,options,resolved)
  expect(out.omissions).toContainEqual({kind:'omission',source:{scope_id:run.id,anchor:run.anchor},code:'hidden-text',count:1,diagnostic_ids:[]})
 })
 it('retains scoped source diagnostics while preserving unrelated text',()=>{
  const {document,resolved}=fixture(),p=document.body.blocks[0]!.paragraph!,run=p.runs.find(r=>r.kind==='drawing')!
  document.unsupported=[{id:'unsafe',code:'UNKNOWN_DRAWING',capability:'drawing',scope_id:run.id,anchor:run.anchor,preservation:'preserve-verbatim',message:'Cannot render'}]
  const out=project(document,options,resolved)
  expect(out.coverage.projected_text_runs).toBeGreaterThan(0)
  expect(out.omissions.some(o=>o.diagnostic_ids.includes('unsafe'))).toBe(true)
  expect(out.source_diagnostics.document).toEqual(document.unsupported)
 })
 it('unknown global or inconsistent diagnostic ownership prevents source text qualification',()=>{
  for(const mismatch of [false,true]){
   const {document,resolved}=fixture(),p=document.body.blocks[0]!.paragraph!
   document.unsupported=[{id:'unsafe',code:'UNKNOWN_SOURCE',capability:'unknown',scope_id:mismatch?p.id:document.document_id,...(mismatch?{anchor:{...p.anchor,path:'/w:document[1]/w:body[1]/w:p[500]'}}:{}),preservation:'preserve-verbatim',message:'Unknown'}]
   const out=project(document,options,resolved)
   expect(out.coverage.projected_text_runs).toBe(0)
   expect(out.omissions.some(o=>o.diagnostic_ids.includes('unsafe'))).toBe(true)
  }
  const {document,resolved}=fixture(),p=document.body.blocks[0]!.paragraph!
  resolved.diagnostics=[{code:'UNKNOWN_STYLE',scope_id:p.id,part_name:document.source.main_part,path:'/w:document[1]/w:body[1]/w:p[500]',severity:'unsupported',preservation:'preserve-verbatim',message:'Unbound source effects'}]
  const result=project(document,options,resolved)
  expect(result.coverage.projected_text_runs).toBe(0)
  expect(result.source_diagnostics.resolved).toEqual(resolved.diagnostics)
 })
 it('refuses malformed contracts, stale layouts and implicit policy',()=>{
  const {document,resolved}=fixture()
  expect(()=>project({...document,revision:null},options,resolved)).toThrow('integrity')
  expect(()=>project(document,options,{...resolved,revision:'stale'})).toThrow('identity')
  expect(()=>project(document,{...options,read_only:false} as never,resolved)).toThrow('Explicit')
  expect(()=>project(document,{...options,extra:true} as never,resolved)).toThrow('Explicit')
  let getterCalls=0
  const accessor={read_only:true,get policy(){getterCalls++;return DOCX_PARTIAL_CONTENT_POLICY}}
  expect(()=>project(document,accessor as never,resolved)).toThrow('Explicit');expect(getterCalls).toBe(0)
  expect(()=>project(document,{...options,[Symbol('hidden')]:true},resolved)).toThrow('Explicit')
  expect(()=>project(document,Object.create(options),resolved)).toThrow('Explicit')
  expect(()=>project(document,new Proxy(options,{getPrototypeOf(){throw new Error('hostile')}}),resolved)).toThrow('Explicit')
 })
 it('reports bounded text omissions without partially slicing a source run',()=>{
  const {document,resolved}=fixture(),run=document.body.blocks[0]!.paragraph!.runs.find(r=>r.kind==='text')!
  run.text='x'.repeat(100001)
  const out=project(document,options,resolved)
  expect(out.omissions.some(o=>o.code==='text-limit'&&o.source.scope_id===run.id)).toBe(true)
 })
 it('reports all body-block truncation units in one explicit source-bound placeholder',()=>{
  const {document}=fixture(),p=document.body.blocks[0]!.paragraph!
  for(let i=0;i<201;i++){const next=structuredClone(p);next.id=`extra:${i}`;next.anchor={...p.anchor,path:`/w:document[1]/w:body[1]/w:p[${i+20}]`};for(const [n,r]of next.runs.entries()){r.id=`extra:${i}:run:${n}`;r.anchor={...r.anchor,path:next.anchor.path+`/w:r[${n+1}]`};if(r.drawing)r.drawing.id=`extra:${i}:drawing:${n}`}document.body.blocks.push({kind:'paragraph',id:next.id,paragraph:next})}
  const out=project(document,options)
  expect(out.coverage.visited_body_blocks).toBe(200)
  expect(out.omissions.find(o=>o.code==='block-limit')).toMatchObject({count:document.body.blocks.length-200,source:{scope_id:document.body.id}})
 })
})
