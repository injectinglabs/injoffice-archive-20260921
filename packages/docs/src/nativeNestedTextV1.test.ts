import {readFileSync} from 'node:fs'
import {it,expect} from 'vitest'
import {createNativeDocxNestedTextInventoryV1 as project,decodeNativeDocxNestedTextEvidenceV1 as decode,type NativeDocxNestedTextEvidenceV1} from './nativeNestedTextV1.js'
import {createNativeDocxPartialContentPreviewV1} from './nativePartialContentV1.js'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
import type {NativeDocxNestedTableOmissionsV1} from './nativePartialNestedTablesV1.js'
import type {NativeDocxPartialTableTextContextV1} from './nativePartialTableTextContextV1.js'
const fixture=()=>JSON.parse(readFileSync(new URL('../../../testdata/docx-native/nested-text-inspection-v1.json',import.meta.url),'utf8')) as {document:NativeDocxDocumentV1;resolved_layout:NativeDocxResolvedLayoutInputV1;nested_table_omissions:NativeDocxNestedTableOmissionsV1;table_text_contexts:NativeDocxPartialTableTextContextV1[];nested_text:NativeDocxNestedTextEvidenceV1[]}
const options={policy:'source-nested-table-text-v1' as const,read_only:true as const}
const run=(f:ReturnType<typeof fixture>,e:unknown=f.nested_text)=>project(f.document,options,f.resolved_layout,f.nested_table_omissions,f.table_text_contexts,e)
it('shows only explicitly requested exact nested text while preserving default omissions and diagnostics',()=>{
 const f=fixture(),before=structuredClone(f),out=run(f)
 expect(JSON.stringify(out.items)).toContain('Inner visible');expect(out.omitted_count).toBe(0)
 expect(out.source_diagnostics).toEqual({document:f.document.unsupported,resolved:f.resolved_layout.diagnostics})
 const old=createNativeDocxPartialContentPreviewV1(f.document,{policy:'source-text-with-omissions-v1',read_only:true},f.resolved_layout,f.nested_table_omissions,f.table_text_contexts)
 expect(JSON.stringify(old.blocks)).not.toContain('Inner visible');expect(old.omissions.some(x=>x.code==='nested-table')).toBe(true)
 expect(project(f.document,options,f.resolved_layout,f.nested_table_omissions,f.table_text_contexts,undefined).items).toEqual([])
 expect(f).toEqual(before)
 expect(()=>project(f.document,{...options,read_only:false} as never,f.resolved_layout,f.nested_table_omissions,f.table_text_contexts,f.nested_text)).toThrow()
})
it('retains global, table, row, owner cell and continuation boundaries',()=>{
 for(const scope of ['global','table','row','cell','continuation','resolved']){
  const f=fixture(),t=f.document.body.blocks[0]!.table!,r=t.rows[0]!,c=r.cells[0]!
  if(scope==='continuation')c.vertical_merge='continue'
  else if(scope==='resolved')f.resolved_layout.diagnostics.push({code:'MISSING_TABLE_STYLE',scope_id:t.id,severity:'unsupported',preservation:'preserve-verbatim',message:'Retained'})
  else{const owner=scope==='table'?t:scope==='row'?r:c;f.document.unsupported.push({id:'other:1',code:'UNMODELED',scope_id:scope==='global'?f.document.document_id:owner.id,anchor:owner.anchor,capability:'source',preservation:'refuse-mutation',message:'Retained'})}
  expect(run(f).items,scope).toEqual([]);expect(run(f).omitted_count).toBe(1)
 }
})
it('rejects stale containment, style, visibility and incomplete resolved evidence',()=>{
 for(const mutate of [
  (e:NativeDocxNestedTextEvidenceV1)=>e.owner.package_sha256='sha256:'+'0'.repeat(64),
  (e:NativeDocxNestedTextEvidenceV1)=>e.owner.cell_id='other',
  (e:NativeDocxNestedTextEvidenceV1)=>e.owner.anchor.xml_sha256='sha256:'+'0'.repeat(64),
  (e:NativeDocxNestedTextEvidenceV1)=>e.style_id='Missing',
  (e:NativeDocxNestedTextEvidenceV1)=>e.paragraphs[0]!.anchor.path+='/w:ins[1]',
  (e:NativeDocxNestedTextEvidenceV1)=>e.paragraphs[0]!.runs[0]!.anchor.end_byte=e.owner.anchor.end_byte+1,
  (e:NativeDocxNestedTextEvidenceV1)=>e.resolved_layout.runs[0]!.properties.hidden=true,
  (e:NativeDocxNestedTextEvidenceV1)=>e.resolved_layout.paragraphs[0]!.paragraph_mark_properties.hidden=true,
  (e:NativeDocxNestedTextEvidenceV1)=>e.resolved_layout.runs=[],
  (e:NativeDocxNestedTextEvidenceV1)=>e.resolved_layout.source_parts.styles_part='other.xml',
  (e:NativeDocxNestedTextEvidenceV1)=>e.resolved_layout.diagnostics.push({code:'MISSING_STYLE_REFERENCE',scope_id:e.paragraphs[0]!.id,severity:'unsupported',preservation:'preserve-verbatim',message:'Retained'}),
  (e:NativeDocxNestedTextEvidenceV1)=>e.paragraphs[0]!.runs[0]!.text='x'.repeat(32769),
  (e:NativeDocxNestedTextEvidenceV1)=>Object.defineProperty(e,'style_id',{get(){throw Error('getter')}}),
 ]){const f=fixture();mutate(f.nested_text[0]!);expect(()=>run(f)).toThrow()}
 const f=fixture();f.nested_text.push(structuredClone(f.nested_text[0]!));expect(()=>run(f)).toThrow()
 expect(()=>decode(f.document,f.resolved_layout,f.nested_table_omissions,undefined,f.nested_text)).toThrow()
})
