import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {createNativeDocxPartialContentPreviewV1,type NativeDocxDocumentV1,type NativeDocxResolvedLayoutInputV1} from '@injoffice/docs/native-docx'
import {NativeDocxPartialText,NativeDocxPartialTextView,NativeDocxEquationList} from './NativeDocxPartialText'
describe('browser-local partial text UI',()=>{
 it('paints indexed radicals with the radicand first and escaped degree text',()=>{
  const html=renderToStaticMarkup(createElement(NativeDocxEquationList,{equations:[{package_sha256:'sha256:'+'a'.repeat(64),paragraph_id:'p:1',diagnostic_id:'equation:1',anchor:{part_name:'word/document.xml',path:'/p/math',start_byte:1,end_byte:2,xml_sha256:'sha256:'+'b'.repeat(64)},status:'supported',tree:{kind:'indexed-radical',children:[{kind:'text',text:'x + 1'},{kind:'text',text:'3<script>'}]}}]}))
  expect(html).toContain('<mroot><mtext>x + 1</mtext><mtext>3&lt;script&gt;</mtext></mroot>')
  expect(html).not.toContain('<script>');expect(html).not.toContain('contenteditable')
 })
 it('uses fixed MathML elements and escapes equation source text',()=>{
  const html=renderToStaticMarkup(createElement(NativeDocxEquationList,{equations:[{package_sha256:'sha256:'+'a'.repeat(64),paragraph_id:'p:1',diagnostic_id:'equation:1',anchor:{part_name:'word/document.xml',path:'/p/math',start_byte:1,end_byte:2,xml_sha256:'sha256:'+'b'.repeat(64)},status:'supported',tree:{kind:'fraction',children:[{kind:'text',text:'<script>bad</script>'},{kind:'radical',children:[{kind:'text',text:'x'}]}]}}]}))
  expect(html).toContain('<mfrac>');expect(html).toContain('<msqrt>');expect(html).toContain('&lt;script&gt;');expect(html).not.toContain('<script>');expect(html).toContain('not Word typography or pagination')
 })
 it('requires explicit request and explains browser-only behavior without upload or edit controls',()=>{
  const html=renderToStaticMarkup(createElement(NativeDocxPartialText,{bytes:new Uint8Array([1]),packageDigest:'sha256:'+'a'.repeat(64)}))
  expect(html).toContain('Show read-only partial text');expect(html).toContain('No file is uploaded')
  expect(html).not.toContain('Read-only partial source text')
 })
 it('renders same-source library table text and authored descriptions as escaped read-only content',()=>{
  const document=JSON.parse(readFileSync(new URL('../../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1
  document.unsupported=[]
  const resolved:NativeDocxResolvedLayoutInputV1={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
  for(const block of document.body.blocks)for(const p of block.paragraph?[block.paragraph]:block.table!.rows.flatMap(r=>r.cells.flatMap(c=>c.paragraphs))){resolved.paragraphs.push({paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}});for(const r of p.runs){resolved.runs.push({run_id:r.id,paragraph_id:p.id,applied_paragraph_styles:[],applied_character_styles:[],properties:{}});if(r.drawing)r.drawing.alt_text='<script>description</script>'}}
  const preview=createNativeDocxPartialContentPreviewV1(document,{policy:'source-text-with-omissions-v1',read_only:true},resolved)
  const html=renderToStaticMarkup(createElement(NativeDocxPartialTextView,{preview}))
  expect(html).toContain('Summary');expect(html).toContain('Source row 1, cell 1');expect(html).toContain('no table layout')
  expect(html).toContain('Authored drawing description');expect(html).toContain('&lt;script&gt;description&lt;/script&gt;')
  expect(html).not.toContain('<script>');expect(html).not.toContain('contenteditable');expect(html).not.toContain('<input')
  const cell=document.body.blocks[1]!.table!.rows[0]!.cells[0]!
  cell.grid_span=2;cell.vertical_merge='restart'
  const owner=renderToStaticMarkup(createElement(NativeDocxPartialTextView,{preview:createNativeDocxPartialContentPreviewV1(document,{policy:'source-text-with-omissions-v1',read_only:true},resolved)}))
  expect(owner).toContain('Summary');expect(owner).toContain('Authored span: 2 grid columns');expect(owner).toContain('Vertical merge starts here — owner text only');expect(owner).toContain('Merge geometry is not reproduced')
  cell.grid_span=1;cell.vertical_merge='continue'
  const continuation=renderToStaticMarkup(createElement(NativeDocxPartialTextView,{preview:createNativeDocxPartialContentPreviewV1(document,{policy:'source-text-with-omissions-v1',read_only:true},resolved)}))
  expect(continuation).not.toContain('Summary');expect(continuation).toContain('Vertical merge continuation — text omitted')
  expect(continuation).toContain('1 grid column.');expect(continuation).not.toContain('1 grid columns')
 })
})
