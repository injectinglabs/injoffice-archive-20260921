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
 function storyFixture(){
  const joined=fixture()
  for(const story of [...joined.document.headers,...joined.document.footers]){
   story.anchor.end_byte=1000
   const p=structuredClone(joined.document.body.blocks[0]!.paragraph!)
   p.id=story.id+':p';p.anchor={...p.anchor,part_name:story.part_name,path:story.anchor.path+'/w:p[1]'}
   p.runs=[p.runs[0]!];const run=p.runs[0]!
   run.id=p.id+':r';run.anchor={...run.anchor,part_name:story.part_name,path:p.anchor.path+'/w:r[1]'};run.text=story.kind+' inventory <script>'
   story.blocks=[{kind:'paragraph',id:p.id,paragraph:p}]
   joined.resolved.paragraphs.push({paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}})
   joined.resolved.runs.push({run_id:run.id,paragraph_id:p.id,applied_paragraph_styles:[],applied_character_styles:[],properties:{}})
  }
  return joined
 }
 it('recovers separately labeled header/footer source paragraphs without selecting active variants or changing source',()=>{
  const {document,resolved}=storyFixture(),before=structuredClone({document,resolved}),out=project(document,options,resolved)
  expect(out.header_footer_stories).toHaveLength(2)
  for(const [index,story]of [...document.headers,...document.footers].entries()){
   expect(out.header_footer_stories![index]).toMatchObject({kind:story.kind,source:{scope_id:story.id,anchor:story.anchor},page_assignment:'not-selected'})
   expect(JSON.stringify(out.header_footer_stories![index]!.blocks)).toContain(story.kind+' inventory <script>')
  }
  expect(JSON.stringify(out.blocks)).not.toContain('inventory <script>')
  expect({document,resolved}).toEqual(before)
  expect(out.pagination).toBe('not-produced')
 })
 it('retains all global/story/paragraph/run blockers, hidden text and missing joins in header/footer inventory',()=>{
  for(const scope of ['global','story','paragraph','run','deleted','hidden','missing']){
   const {document,resolved}=storyFixture(),story=document.headers[0]!,p=story.blocks[0]!.paragraph!,run=p.runs[0]!
   if(scope==='hidden')resolved.runs.find(r=>r.run_id===run.id)!.properties.hidden=true
   else if(scope==='missing')resolved.runs=resolved.runs.filter(r=>r.run_id!==run.id)
   else{
    const owner=scope==='story'?story:scope==='run'?run:p
    document.unsupported=[{id:'unsafe',code:scope==='deleted'?'UNMODELED_PARAGRAPH_MARKUP':'UNMODELED',scope_id:scope==='global'?document.document_id:owner.id,anchor:owner.anchor,capability:'source',preservation:'preserve-verbatim',message:'Retained'}]
   }
   const out=project(document,options,resolved)
   expect(JSON.stringify(out.header_footer_stories![0]!.blocks)).not.toContain('header inventory')
   expect(out.source_diagnostics.document).toEqual(document.unsupported)
   if(scope==='global')expect(JSON.stringify(out.header_footer_stories)).not.toContain('footer inventory')
  }
  const {document}=storyFixture()
  expect(JSON.stringify(project(document,options).header_footer_stories)).not.toContain('inventory <script>')
 })
 it('shares text budgets with the body and across header/footer stories',()=>{
  const {document,resolved}=storyFixture()
  document.body.blocks[0]!.paragraph!.runs[0]!.text='x'.repeat(100_000)
  const out=project(document,options,resolved)
  expect(JSON.stringify(out.header_footer_stories)).not.toContain('inventory <script>')
  expect(out.header_footer_stories!.every(s=>JSON.stringify(s.blocks).includes('text-limit'))).toBe(true)
 })
 it('requires exact header/footer source parts and preserves resolved blockers',()=>{
  const {document,resolved}=storyFixture(),story=document.headers[0]!,p=story.blocks[0]!.paragraph!
  p.runs[0]!.anchor.part_name='word/footer1.xml'
  expect(()=>project(document,options,resolved)).toThrow('Invalid native document')
  p.runs[0]!.anchor.part_name=story.part_name
  resolved.diagnostics=[{code:'UNMODELED',scope_id:p.id,part_name:story.part_name,path:p.anchor.path,severity:'unsupported',preservation:'preserve-verbatim',message:'Retained'}]
  const out=project(document,options,resolved)
  expect(JSON.stringify(out.header_footer_stories![0]!.blocks)).not.toContain('header inventory')
  expect(out.source_diagnostics.resolved).toEqual(resolved.diagnostics)
  resolved.diagnostics[0]!.scope_id=document.document_id
  expect(JSON.stringify(project(document,options,resolved).header_footer_stories)).not.toContain('inventory <script>')
 })
 it('bounds the combined header/footer block inventory and leaves tables opaque',()=>{
  const {document}=storyFixture()
  for(const [index,story]of [...document.headers,...document.footers].entries()){
   const original=story.blocks[0]!.paragraph!
   story.blocks=Array.from({length:index===0?150:51},(_,i)=>{
    const p=structuredClone(original);p.id+=`:${i}`;p.anchor.path=story.anchor.path+`/w:p[${i+1}]`
    p.runs[0]!.id+=`:${i}`;p.runs[0]!.anchor.path=p.anchor.path+'/w:r[1]'
    return {kind:'paragraph' as const,id:p.id,paragraph:p}
   })
  }
  const out=project(document,options)
  expect(out.header_footer_stories![0]!.blocks).toHaveLength(150)
  expect(out.header_footer_stories![1]!.blocks).toHaveLength(51)
  expect(out.header_footer_stories![1]!.blocks[50]).toMatchObject({code:'block-limit',count:1})
  const {document:withTable,resolved}=storyFixture(),story=withTable.headers[0]!,table=withTable.body.blocks.splice(1,1)[0]!
  const rewrite=(anchor:typeof story.anchor)=>{anchor.part_name=story.part_name;anchor.path=anchor.path.replace('/w:document[1]/w:body[1]',story.anchor.path)}
  rewrite(table.table!.anchor);story.anchor.end_byte=10000
  for(const row of table.table!.rows){rewrite(row.anchor);for(const cell of row.cells){rewrite(cell.anchor);for(const p of cell.paragraphs){rewrite(p.anchor);for(const r of p.runs)rewrite(r.anchor)}}}
  story.blocks.push(table)
  const result=project(withTable,options,resolved)
  expect(result.header_footer_stories![0]!.blocks[1]).toMatchObject({code:'table',source:{scope_id:table.id}})
  expect(JSON.stringify(result.header_footer_stories)).not.toContain('Summary')
 })
 it('recovers outer paragraph text around an exact nested-table omission without changing strict source',()=>{
  const {document,resolved}=fixture(),table=document.body.blocks[1]!.table!,cell=table.rows[0]!.cells[0]!
  const anchor={...cell.anchor,path:cell.anchor.path+'/w:tbl[1]',start_byte:1401,end_byte:1450}
  const diagnostic={id:'nested:1',code:'NESTED_TABLE_OR_CELL_MARKUP',scope_id:table.id,anchor,capability:'table-structure',preservation:'refuse-mutation' as const,message:'Nested source remains opaque'}
  document.unsupported=[diagnostic]
  const evidence={items:[{package_sha256:document.source.package_sha256,part_sha256:'sha256:'+'a'.repeat(64),table_id:table.id,cell_id:cell.id,diagnostic_id:diagnostic.id,anchor}],omitted_count:0}
  const before=structuredClone({document,resolved,evidence}),out=project(document,options,resolved,evidence)
  expect(JSON.stringify(out.blocks)).toContain('Summary');expect(out.omissions).toContainEqual({kind:'omission',source:{scope_id:diagnostic.id,anchor},code:'nested-table',count:1,diagnostic_ids:[diagnostic.id]})
  expect(out.source_diagnostics.document).toEqual([diagnostic]);expect(out.nested_table_omissions).toEqual(evidence);expect({document,resolved,evidence}).toEqual(before)
  expect(JSON.stringify(project(document,options,resolved).blocks)).not.toContain('Summary')
  cell.grid_span=2;cell.vertical_merge='restart'
  expect(JSON.stringify(project(document,options,resolved,evidence).blocks)).toContain('Summary')
  cell.vertical_merge='continue'
  expect(JSON.stringify(project(document,options,resolved,evidence).blocks)).not.toContain('Summary')
  cell.vertical_merge='restart'
  const run=cell.paragraphs[0]!.runs[0]!;resolved.runs.find(r=>r.run_id===run.id)!.properties.hidden=true
  expect(JSON.stringify(project(document,options,resolved,evidence).blocks)).not.toContain('Summary')
  delete resolved.runs.find(r=>r.run_id===run.id)!.properties.hidden
  document.unsupported.push({...diagnostic,id:'unknown',code:'UNMODELED_CELL_PROPERTY'})
  expect(JSON.stringify(project(document,options,resolved,evidence).blocks)).not.toContain('Summary')
 })
 it('rejects forged nested evidence and hostile dense-array inputs without invoking getters',()=>{
  const {document,resolved}=fixture(),table=document.body.blocks[1]!.table!,cell=table.rows[0]!.cells[0]!,anchor={...cell.anchor,path:cell.anchor.path+'/w:tbl[1]',start_byte:1401,end_byte:1450}
  document.unsupported=[{id:'nested:1',code:'NESTED_TABLE_OR_CELL_MARKUP',scope_id:table.id,anchor,capability:'table-structure',preservation:'refuse-mutation',message:'Opaque'}]
  const fact={package_sha256:document.source.package_sha256,part_sha256:'sha256:'+'a'.repeat(64),table_id:table.id,cell_id:cell.id,diagnostic_id:'nested:1',anchor}
  for(const change of [{package_sha256:'sha256:'+'0'.repeat(64)},{cell_id:'wrong'},{part_sha256:'bad'},{anchor:{...anchor,path:anchor.path+'/w:p[1]'}},{anchor:{...anchor,start_byte:1100}},{anchor:{...anchor,xml_sha256:'sha256:'+'0'.repeat(64)}}])expect(()=>project(document,options,resolved,{items:[{...fact,...change}],omitted_count:0})).toThrow()
  let called=0;const getter=Object.defineProperty([],0,{get(){called++;return fact}})
  for(const items of [getter,new Array(1),[fact,fact],Object.assign([fact],{extra:1}),Object.assign([fact],{[Symbol('x')]:1}),Array(65).fill(fact)])expect(()=>project(document,options,resolved,{items,omitted_count:0})).toThrow()
  expect(called).toBe(0)
  const second={...fact,diagnostic_id:'nested:2',part_sha256:'sha256:'+'b'.repeat(64),anchor:{...anchor,path:cell.anchor.path+'/w:tbl[2]',start_byte:1451,end_byte:1490}}
  document.unsupported.push({...document.unsupported[0]!,id:second.diagnostic_id,anchor:second.anchor})
  expect(()=>project(document,options,resolved,{items:[fact,second],omitted_count:0})).toThrow('source identity')
  document.unsupported.pop()
  const out=project(document,options,resolved,{items:[fact],omitted_count:3})
  expect(out.omissions.some(o=>o.code==='nested-table-limit'&&o.count===3)).toBe(true)
 })
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
