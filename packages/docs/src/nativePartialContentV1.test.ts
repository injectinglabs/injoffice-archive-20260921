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
