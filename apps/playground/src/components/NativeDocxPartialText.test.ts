import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {createNativeDocxNestedTextInventoryV1,createNativeDocxPartialContentPreviewV1,type NativeDocxDocumentV1,type NativeDocxResolvedLayoutInputV1} from '@injoffice/docs/native-docx'
import {NativeDocxPartialText,NativeDocxPartialTextView,NativeDocxEquationList,NativeDocxReviewInventoryView,NativeDocxNestedTextInventoryView} from './NativeDocxPartialText'
describe('browser-local partial text UI',()=>{
 it('labels nested source text separately, retains omissions and escapes authored text',()=>{
  const f=JSON.parse(readFileSync(new URL('../../../../testdata/docx-native/nested-text-inspection-v1.json',import.meta.url),'utf8'))
  const inventory=createNativeDocxNestedTextInventoryV1(f.document,{policy:'source-nested-table-text-v1',read_only:true},f.resolved_layout,f.nested_table_omissions,f.table_text_contexts,f.nested_text)
  const segment=inventory.items[0]!.paragraphs[0]!.segments[0]!;if(segment.kind==='text')segment.text='<script>nested</script>'
  const html=renderToStaticMarkup(createElement(NativeDocxNestedTextInventoryView,{inventory}))
  expect(html).toContain('&lt;script&gt;nested&lt;/script&gt;');expect(html).toContain('Original nested-table omissions and diagnostics remain');expect(html).toContain('does not reconstruct table geometry')
  expect(html).not.toContain('<script>');expect(html).not.toContain('<input');expect(html).not.toContain('<button');expect(html).not.toContain('contenteditable')
 })
 it('labels review kinds, metadata, omissions and retained diagnostic codes as escaped read-only text',()=>{
  const anchor={part_name:'word/document.xml',path:'/w:document[1]/w:body[1]/w:p[1]/w:ins[1]',start_byte:1,end_byte:2,xml_sha256:'sha256:'+'b'.repeat(64)}
  const html=renderToStaticMarkup(createElement(NativeDocxReviewInventoryView,{review:{policy:'source-review-inventory-v1',read_only:true,source:{document_id:'d',revision:'r',package_sha256:'sha256:'+'a'.repeat(64)},items:[{package_sha256:'sha256:'+'a'.repeat(64),paragraph_id:'p',diagnostic_id:'diag',anchor,kind:'insertion',revision_id:'1',author:'<script>Author</script>',created_at:'<date>',run_ids:['r'],segments:[{kind:'text',source:{scope_id:'r',anchor},text:'<script>Inserted</script>'}],text_status:'qualified-insertion',retained_diagnostic_ids:['diag']}],omitted_count:1,source_diagnostics:{document:[{id:'diag',code:'WRAPPED_RUN_MARKUP',scope_id:'p',anchor,capability:'run-structure',preservation:'refuse-mutation',message:'Retained'}],resolved:[]}}}))
  expect(html).toContain('insertion 1');expect(html).toContain('&lt;script&gt;Author&lt;/script&gt;');expect(html).toContain('&lt;script&gt;Inserted&lt;/script&gt;');expect(html).toContain('WRAPPED_RUN_MARKUP');expect(html).toContain('Review inventory omissions: 1');expect(html).toContain('Deleted and moved text stays omitted')
  expect(html).not.toContain('<script>');expect(html).not.toContain('contenteditable');expect(html).not.toContain('<input');expect(html).not.toContain('<button')
 })
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
  expect(html).toContain('Show read-only partial text');expect(html).toContain('Show partial text with comments');expect(html).toContain('Inspect tracked-change source');expect(html).toContain('No file is uploaded')
  expect(html).not.toContain('Read-only partial source text')
 })
 it('renders same-source library table text and authored descriptions as escaped read-only content',()=>{
  const document=JSON.parse(readFileSync(new URL('../../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1
  document.unsupported=[]
  const resolved:NativeDocxResolvedLayoutInputV1={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
  for(const block of document.body.blocks)for(const p of block.paragraph?[block.paragraph]:block.table!.rows.flatMap(r=>r.cells.flatMap(c=>c.paragraphs))){resolved.paragraphs.push({paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}});for(const r of p.runs){resolved.runs.push({run_id:r.id,paragraph_id:p.id,applied_paragraph_styles:[],applied_character_styles:[],properties:{}});if(r.drawing)r.drawing.alt_text='<script>description</script>'}}
  const preview=createNativeDocxPartialContentPreviewV1(document,{policy:'source-text-with-omissions-v1',read_only:true},resolved)
  const anchor=document.headers[0]!.anchor
  preview.header_footer_stories=[{kind:'header',source:{scope_id:'source:header',anchor:{...anchor,part_name:'word/<script>header.xml'}},page_assignment:'not-selected',blocks:[{kind:'paragraph',source:{scope_id:'source:p',anchor},segments:[{kind:'text',source:{scope_id:'source:r',anchor},text:'Header <script>source</script>'}]}]}]
  preview.comment_inventory={policy:'source-comment-inventory-v1',stories:[{kind:'comment',source:{scope_id:'comment:story',anchor},comment_source:{scope_id:'comment:1',anchor},native_comment_id:'1',author:'<script>Author</script>',created_at:'<script>date</script>',range_assignment:'not-reconstructed',blocks:[{kind:'paragraph',source:{scope_id:'comment:p',anchor},segments:[{kind:'text',source:{scope_id:'comment:r',anchor},text:'Comment <script>body</script>'}]}]}]}
  preview.table_text_contexts=[{package_sha256:document.source.package_sha256,table_id:'table:source',look_diagnostic_id:'look:1',look_anchor:anchor,styles_part:'word/styles.xml',styles_sha256:'sha256:'+'a'.repeat(64),style_chain:[],resolved_diagnostics:[]}]
  const html=renderToStaticMarkup(createElement(NativeDocxPartialTextView,{preview}))
  expect(html).toContain('1 source-qualified geometry-only table style contexts allow plain text recovery');expect(html).toContain('Conditional styles remain unsupported')
  expect(html).toContain('Read-only comment source inventory');expect(html).toContain('Comment &lt;script&gt;body&lt;/script&gt;');expect(html).toContain('&lt;script&gt;Author&lt;/script&gt;');expect(html).toContain('Stored date: &lt;script&gt;date&lt;/script&gt;');expect(html).toContain('cannot accept or reject changes')
  expect(html).toContain('Header and footer source inventory');expect(html).toContain('No active first, even or default variant is selected')
  expect(html).toContain('word/&lt;script&gt;header.xml');expect(html).toContain('Header &lt;script&gt;source&lt;/script&gt;')
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
