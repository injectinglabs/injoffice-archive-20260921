import { readFileSync } from 'node:fs'
import {createHash} from 'node:crypto'
import { describe, expect, it } from 'vitest'
import {
  NATIVE_WASM_WORKER_PROTOCOL,
  NATIVE_WASM_WORKER_VERSION,
  type NativeWasmMessageEvent,
  type NativeWasmWorker,
  type NativeWasmWorkerErrorEvent,
  type NativeWasmWorkerRequest,
  type NativeWasmWorkerResponse,
} from '@injoffice/native-runtime'
import {
  OFFICE_MUTATION_PROTOCOL,
  OFFICE_MUTATION_VERSION,
  NativeDocxValidationError,
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
  type NativeDocxOfficeMutationEnvelopeV1,
  type NativeDocxTextMutationPayloadV1,
} from '@injoffice/docs'
import {
  DOCX_WASM_NATIVE_MAX_PACKAGE_BYTES,
  createDocxWasmClient,
  resolveDocxWasmAssetUrls,
} from './index'

const fixtureRecord = JSON.parse(readFileSync(
  new URL('../../../go/officecompat/corpus/generated/expected/docx-strict-relocated.json', import.meta.url),
  'utf8',
)) as { native: unknown }
const fixtureDecoded = decodeNativeDocxDocument(fixtureRecord.native)
if (!fixtureDecoded.ok) throw new Error('DOCX WASM test fixture must be valid')
const fixtureDocument = fixtureDecoded.value
const fixtureRun = fixtureDocument.body.blocks[0].paragraph!.runs[0]
const fixtureParagraph = fixtureDocument.body.blocks[0].paragraph!
const multiRunRecord = JSON.parse(readFileSync(
  new URL('../../../go/officecompat/corpus/generated/expected/docx-inline-png-page-paint.json', import.meta.url),
  'utf8',
)) as { native: unknown }
const multiRunDecoded = decodeNativeDocxDocument(multiRunRecord.native)
if (!multiRunDecoded.ok) throw new Error('DOCX WASM multi-run fixture must be valid')
const multiRunDocument = multiRunDecoded.value
const multiRunParagraph = multiRunDocument.body.blocks.map((block) => block.paragraph).find((paragraph) => paragraph && paragraph.runs.length > 1)!

const envelope = (document = fixtureDocument): NativeDocxOfficeMutationEnvelopeV1 & { payload: NativeDocxTextMutationPayloadV1 } => ({
  protocol: OFFICE_MUTATION_PROTOCOL,
  version: OFFICE_MUTATION_VERSION,
  format: 'docx',
  mutation_id: 'save-1',
  expected_revision: document.source.package_sha256,
  payload: {
    mutations: [{
      target_kind: 'run',
      target_id: fixtureRun.id,
      expected_xml_sha256: fixtureRun.anchor.xml_sha256,
      text: 'Strict native',
    }],
  },
})

class FakeWorker implements NativeWasmWorker {
  readonly requests: NativeWasmWorkerRequest[] = []
  terminated = false
  private readonly messageListeners = new Set<(event: NativeWasmMessageEvent) => void>()
  private readonly errorListeners = new Set<(event: NativeWasmWorkerErrorEvent) => void>()

  constructor(private readonly extractJson = JSON.stringify(fixtureDocument)) {}

  postMessage(value: unknown, _transfer: ArrayBuffer[]): void {
    const request = value as NativeWasmWorkerRequest
    this.requests.push(request)
    if (request.op === 'init') this.respond(success(request))
    else if (request.op === 'extract'||request.op==='inspect') this.respond(success(request, { contractJson: this.extractJson }))
    else this.respond(success(request, { bytes: new Uint8Array([4, 5, 6]).buffer }))
  }

  terminate(): void { this.terminated = true }

  addEventListener(type: 'message' | 'error', listener: ((event: NativeWasmMessageEvent) => void) | ((event: NativeWasmWorkerErrorEvent) => void)): void {
    if (type === 'message') this.messageListeners.add(listener as (event: NativeWasmMessageEvent) => void)
    else this.errorListeners.add(listener as (event: NativeWasmWorkerErrorEvent) => void)
  }

  removeEventListener(type: 'message' | 'error', listener: ((event: NativeWasmMessageEvent) => void) | ((event: NativeWasmWorkerErrorEvent) => void)): void {
    if (type === 'message') this.messageListeners.delete(listener as (event: NativeWasmMessageEvent) => void)
    else this.errorListeners.delete(listener as (event: NativeWasmWorkerErrorEvent) => void)
  }

  private respond(response: NativeWasmWorkerResponse): void {
    queueMicrotask(() => {
      for (const listener of this.messageListeners) listener({ data: response })
    })
  }
}

const success = (request: NativeWasmWorkerRequest, result?: unknown): NativeWasmWorkerResponse => ({
  protocol: NATIVE_WASM_WORKER_PROTOCOL,
  version: NATIVE_WASM_WORKER_VERSION,
  id: request.id,
  format: request.format,
  op: request.op,
  ok: true,
  ...(request.op === 'init' ? {} : { result }),
} as NativeWasmWorkerResponse)

describe('DOCX WASM package client', () => {
  it('joins nested text producer evidence and refuses hidden auxiliary resolution',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex')
    const envelope=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/nested-text-inspection-v1.json',import.meta.url),'utf8'))
    envelope.package_sha256=hash;envelope.document.source.package_sha256=hash
    for(const c of envelope.table_text_contexts)c.package_sha256=hash
    for(const n of envelope.nested_table_omissions.items)n.package_sha256=hash
    for(const n of envelope.nested_text)n.owner.package_sha256=hash
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    expect((await client.inspectPartialContent(bytes)).nested_text?.[0]?.paragraphs[0]?.runs[0]?.text).toBe('Inner visible');client.terminate()
    envelope.nested_text[0].resolved_layout.runs[0].properties.hidden=true
    const badWorker=new FakeWorker(JSON.stringify(envelope)),bad=createDocxWasmClient({workerFactory:()=>badWorker})
    await expect(bad.inspectPartialContent(bytes)).rejects.toThrow();expect(badWorker.terminated).toBe(true)
  })
  it('validates optional table text contexts against source look and style part hashes',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex'),document=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1,table=document.body.blocks[1]!.table!
    document.source.package_sha256=hash
    const anchor={...table.anchor,path:table.anchor.path+'/w:tblPr[1]/w:tblLook[1]',start_byte:table.anchor.start_byte+1,end_byte:table.anchor.start_byte+2}
    document.unsupported=[{id:'look:1',code:'UNMODELED_TABLE_PROPERTY',scope_id:table.id,anchor,capability:'table-properties',preservation:'refuse-mutation',message:'Look'}]
    const resolved_layout={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part,styles_part:'word/styles.xml'},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
    const context={package_sha256:hash,table_id:table.id,look_diagnostic_id:'look:1',look_anchor:anchor,styles_part:'word/styles.xml',styles_sha256:document.passthrough_parts.find(p=>p.part_name==='word/styles.xml')!.sha256,style_chain:[{style_id:table.table_style_id,anchor:{...anchor,part_name:'word/styles.xml',path:'/w:styles[1]/w:style[1]',start_byte:1,end_byte:400}}],resolved_diagnostics:[]}
    const envelope={protocol:'injoffice.docx.partial-source',version:1,package_sha256:hash,document,resolved_layout,table_text_contexts:[context]}
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    expect((await client.inspectPartialContent(bytes)).table_text_contexts).toEqual([context]);client.terminate()
    context.styles_sha256='sha256:'+'0'.repeat(64)
    const badWorker=new FakeWorker(JSON.stringify(envelope)),bad=createDocxWasmClient({workerFactory:()=>badWorker})
    await expect(bad.inspectPartialContent(bytes)).rejects.toThrow();expect(badWorker.terminated).toBe(true)
  })
  it('validates textbox source evidence and terminates on forged anchors',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex'),document=structuredClone(fixtureDocument)
    document.source.package_sha256=hash
    const p=document.body.blocks[0]!.paragraph!,anchor={...p.anchor,path:p.anchor.path+'/w:r[1]/w:pict[1]',start_byte:200,end_byte:220}
    document.unsupported=[{id:'textbox:1',code:'UNMODELED_DRAWING',capability:'drawings',scope_id:p.id,anchor,preservation:'refuse-mutation',message:'Drawing'}]
    const fact={package_sha256:hash,part_sha256:'sha256:'+'a'.repeat(64),paragraph_id:p.id,diagnostic_id:'textbox:1',anchor:{...anchor},kind:'vml',status:'supported',paragraphs:['Literal <script>text</script>'],reason:''}
    const resolved_layout={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
    const envelope={protocol:'injoffice.docx.partial-source',version:1,package_sha256:hash,document,resolved_layout,textbox_inventory:{items:[fact],omitted_count:0}}
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    expect((await client.inspectPartialContent(bytes)).textbox_inventory).toEqual(envelope.textbox_inventory);client.terminate()
    fact.anchor.xml_sha256='sha256:'+'f'.repeat(64)
    const badWorker=new FakeWorker(JSON.stringify(envelope)),bad=createDocxWasmClient({workerFactory:()=>badWorker})
    await expect(bad.inspectPartialContent(bytes)).rejects.toThrow('does not join');expect(badWorker.terminated).toBe(true)
  })
  it('validates optional rectangle geometry and refuses forged dimensions',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex'),document=structuredClone(fixtureDocument)
    document.source.package_sha256=hash
    const p=document.body.blocks[0]!.paragraph!,anchor={...p.anchor,path:p.anchor.path+'/w:r[1]/w:drawing[1]',start_byte:200,end_byte:220}
    document.unsupported=[{id:'geometry:1',code:'PICTURE_GRAPHIC_REQUIRED',capability:'drawings',scope_id:p.id,anchor,preservation:'refuse-mutation',message:'Drawing'}]
    const owner={package_sha256:hash,part_sha256:'sha256:'+'a'.repeat(64),paragraph_id:p.id,diagnostic_id:'geometry:1',anchor:{...anchor},kind:'drawingml',status:'supported',paragraphs:['Rectangle'],reason:''}
    const geometry={width_emu:2743200,height_emu:914400,insets_emu:[91440,91440,91440,91440],fill_rgb:'FFF2CC',line_rgb:'none',line_width_emu:0,font_family:'DejaVu Sans',font_size_half_points:24,text_rgb:'000000'}
    const resolved_layout={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
    const envelope={protocol:'injoffice.docx.partial-source',version:1,package_sha256:hash,document,resolved_layout,textbox_geometry:{items:[{owner,geometry}],omitted_count:0}}
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    expect((await client.inspectPartialContent(bytes)).textbox_geometry).toEqual(envelope.textbox_geometry);client.terminate()
    geometry.width_emu++
    const badWorker=new FakeWorker(JSON.stringify(envelope)),bad=createDocxWasmClient({workerFactory:()=>badWorker})
    await expect(bad.inspectPartialContent(bytes)).rejects.toThrow('exactly representable');expect(badWorker.terminated).toBe(true)
  })
  it('validates source-bound review metadata and rejects deletion text evidence',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex'),document=structuredClone(fixtureDocument)
    document.source.package_sha256=hash
    const p=document.body.blocks[0]!.paragraph!,anchor={...p.anchor,path:p.anchor.path+'/w:del[1]',start_byte:200,end_byte:220}
    document.unsupported=[{id:'review:1',code:'UNMODELED_PARAGRAPH_CONTENT',capability:'run-structure',scope_id:p.id,anchor,preservation:'refuse-mutation',message:'Deletion'}]
    const fact={package_sha256:hash,paragraph_id:p.id,diagnostic_id:'review:1',anchor,kind:'deletion',revision_id:'7',author:'Author',created_at:'',run_ids:[] as string[]}
    const resolved_layout={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
    const envelope={protocol:'injoffice.docx.partial-source',version:1,package_sha256:hash,document,resolved_layout,review_changes:{items:[fact],omitted_count:0}}
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    expect((await client.inspectPartialContent(bytes)).review_changes).toEqual(envelope.review_changes);client.terminate()
    fact.run_ids=[p.runs[0]!.id]
    const badWorker=new FakeWorker(JSON.stringify(envelope)),bad=createDocxWasmClient({workerFactory:()=>badWorker})
    await expect(bad.inspectPartialContent(bytes)).rejects.toThrow('Only insertion');expect(badWorker.terminated).toBe(true)
  })
  it('validates optional nested omission evidence before returning it',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex'),document=JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8')) as NativeDocxDocumentV1
    document.source.package_sha256=hash
    const table=document.body.blocks[1]!.table!,cell=table.rows[0]!.cells[0]!,anchor={...cell.anchor,path:cell.anchor.path+'/w:tbl[1]',start_byte:1401,end_byte:1450}
    document.unsupported=[{id:'nested:1',code:'NESTED_TABLE_OR_CELL_MARKUP',scope_id:table.id,anchor,capability:'table-structure',preservation:'refuse-mutation',message:'Opaque'}]
    const fact={package_sha256:hash,part_sha256:'sha256:'+'a'.repeat(64),table_id:table.id,cell_id:cell.id,diagnostic_id:'nested:1',anchor}
    const resolved_layout={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
    const envelope={protocol:'injoffice.docx.partial-source',version:1,package_sha256:hash,document,resolved_layout,nested_table_omissions:{items:[fact],omitted_count:0}}
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    expect((await client.inspectPartialContent(bytes)).nested_table_omissions).toEqual(envelope.nested_table_omissions);client.terminate()
    const badWorker=new FakeWorker(JSON.stringify({...envelope,nested_table_omissions:{items:[{...fact,cell_id:'wrong'}],omitted_count:0}})),bad=createDocxWasmClient({workerFactory:()=>badWorker})
    await expect(bad.inspectPartialContent(bytes)).rejects.toThrow();expect(badWorker.terminated).toBe(true)
  })
  it('validates optional equation evidence before returning browser math data',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex'),document=structuredClone(fixtureDocument)
    document.source.package_sha256=hash
    const p=document.body.blocks[0]!.paragraph!,anchor={...p.anchor,path:p.anchor.path+'/ns12345678:oMath[1]',start_byte:p.anchor.start_byte+1,end_byte:p.anchor.end_byte-1}
    document.unsupported=[{id:'equation:1',code:'UNMODELED_PARAGRAPH_CONTENT',scope_id:p.id,anchor,capability:'run-structure',preservation:'preserve-verbatim',message:'Equation'}]
    const resolved_layout={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[{paragraph_id:p.id,applied_styles:[],properties:{},paragraph_mark_properties:{}}],runs:[],tables:[],fonts:[],diagnostics:[]}
    const equation={package_sha256:hash,paragraph_id:p.id,anchor,diagnostic_id:'equation:1',status:'supported',tree:{kind:'text',text:'x'}}
    const envelope={protocol:'injoffice.docx.partial-source',version:1,package_sha256:hash,document,resolved_layout,equations:[equation]}
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    expect((await client.inspectPartialContent(bytes)).equations).toEqual([equation]);client.terminate()
    const section=document.sections[0]!,sectionAnchor={...section.anchor,path:section.anchor.path+'/w:textDirection[1]'}
    document.unsupported.push({id:'section:direction',code:'UNMODELED_SECTION_PROPERTY',scope_id:section.id,anchor:sectionAnchor,capability:'sections',preservation:'refuse-mutation',message:'Preserved direction'})
    const notice={kind:'horizontal-section',package_sha256:hash,part_sha256:'sha256:'+'a'.repeat(64),anchor:sectionAnchor,diagnostic_origin:'document',diagnostic_id:'section:direction',code:'UNMODELED_SECTION_PROPERTY',scope_id:section.id,value:'lrTb'}
    const noticeWorker=new FakeWorker(JSON.stringify({...envelope,equation_context_notices:[notice]})),withNotice=createDocxWasmClient({workerFactory:()=>noticeWorker})
    const joined=await withNotice.inspectPartialContent(bytes);expect(joined.equations?.[0]?.status).toBe('supported');expect(joined.equation_context_notices).toEqual([notice]);expect(joined.document.unsupported).toEqual(document.unsupported);withNotice.terminate()
    const badWorker=new FakeWorker(JSON.stringify({...envelope,equations:[{...equation,diagnostic_id:'wrong'}]})),bad=createDocxWasmClient({workerFactory:()=>badWorker})
    await expect(bad.inspectPartialContent(bytes)).rejects.toThrow();expect(badWorker.terminated).toBe(true)
  })
  it('inspects same-byte source and layout with an independent package digest join',async()=>{
    const bytes=new Uint8Array([1,2,3]),hash='sha256:'+createHash('sha256').update(bytes).digest('hex'),document=structuredClone(fixtureDocument)
    document.source.package_sha256=hash
    const layout={protocol:'injoffice.docx.resolved-layout',version:1,document_id:document.document_id,revision:document.revision,source_parts:{main_part:document.source.main_part},paragraphs:[],runs:[],tables:[],fonts:[],diagnostics:[]}
    const envelope={protocol:'injoffice.docx.partial-source',version:1,package_sha256:hash,document,resolved_layout:layout}
    const worker=new FakeWorker(JSON.stringify(envelope)),client=createDocxWasmClient({workerFactory:()=>worker})
    const promise=client.inspectPartialContent(bytes);bytes[0]=99
    expect(await promise).toEqual({document,resolved_layout:layout})
    expect(worker.requests.map(r=>r.op)).toEqual(['init','inspect'])
    expect(bytes[0]).toBe(99);client.terminate()
    for(const invalid of [{...envelope,package_sha256:'sha256:'+'0'.repeat(64)},{...envelope,resolved_layout:{...layout,revision:'stale'}},{...envelope,extra:true},'not-json']){
      const badWorker=new FakeWorker(typeof invalid==='string'?invalid:JSON.stringify(invalid)),bad=createDocxWasmClient({workerFactory:()=>badWorker})
      await expect(bad.inspectPartialContent(new Uint8Array([1,2,3]))).rejects.toThrow();expect(badWorker.terminated).toBe(true)
    }
    const abort=new AbortController(),cancelWorker=new FakeWorker(JSON.stringify(envelope)),cancel=createDocxWasmClient({workerFactory:()=>cancelWorker})
    const pending=cancel.inspectPartialContent(new Uint8Array([1,2,3]),{signal:abort.signal});abort.abort()
    await expect(pending).rejects.toMatchObject({name:'AbortError'});cancel.terminate()
    const largeWorker=new FakeWorker('x'.repeat(16*1024*1024+1)),large=createDocxWasmClient({workerFactory:()=>largeWorker})
    await expect(large.inspectPartialContent(new Uint8Array([1]))).rejects.toThrow('response budget');expect(largeWorker.terminated).toBe(true)
    const smallWorker=new FakeWorker(),small=createDocxWasmClient({workerFactory:()=>smallWorker,maxPackageBytes:1})
    await expect(small.inspectPartialContent(new Uint8Array([1,2]))).rejects.toThrow();expect(smallWorker.requests).toHaveLength(0);small.terminate()
  })
  it('resolves package-relative defaults and exact explicit overrides', () => {
    const defaults = resolveDocxWasmAssetUrls()
    expect(defaults.workerUrl).toMatch(/\/docxnative\.worker\.js$/)
    expect(defaults.wasmUrl).toMatch(/\/docxnative\.wasm$/)
    expect(defaults.goRuntimeUrl).toMatch(/\/wasm_exec\.js$/)
    expect(resolveDocxWasmAssetUrls({
      workerUrl: new URL('https://cdn.example/worker.js'),
      wasmUrl: 'https://cdn.example/engine.wasm',
      goRuntimeUrl: 'https://cdn.example/go.js',
    })).toEqual({
      workerUrl: 'https://cdn.example/worker.js',
      wasmUrl: 'https://cdn.example/engine.wasm',
      goRuntimeUrl: 'https://cdn.example/go.js',
    })
  })

  it('is lazy, validates extraction, and sends the exact CAS-bound payload', async () => {
    const worker = new FakeWorker()
    let calls = 0
    const client = createDocxWasmClient({
      workerUrl: '/worker.js', wasmUrl: '/engine.wasm', goRuntimeUrl: '/go.js',
      workerFactory: (url) => { calls++; expect(url).toBe('/worker.js'); return worker },
    })
    expect(calls).toBe(0)
    const document = await client.extract(new Uint8Array([1, 2, 3]))
    expect(document.version).toBe(1)
    expect(calls).toBe(1)
    await expect(client.apply(new Uint8Array([1]), document, envelope(document))).resolves.toEqual(new Uint8Array([4, 5, 6]))
    expect(worker.requests.map(({ op }) => op)).toEqual(['init', 'extract', 'apply'])
    expect(worker.requests[2]).toMatchObject({
      expectedRevision: document.source.package_sha256,
      payload: JSON.stringify(envelope(document).payload),
    })
  })

  it('validates and forwards body paragraph insertion without enabling other block payloads', async () => {
    const worker = new FakeWorker()
    const client = createDocxWasmClient({workerFactory: () => worker})
    const document = structuredClone(fixtureDocument)
    const paragraph = document.body.blocks.find(block => block.paragraph)!.paragraph!
    if (!paragraph.edit_policy.allowed_operations.includes('block.insert_after')) paragraph.edit_policy.allowed_operations.push('block.insert_after')
    const mutation = {target_kind: 'paragraph' as const, target_id: paragraph.id, expected_xml_sha256: paragraph.anchor.xml_sha256, operation: 'block.insert_after' as const, text: '' as const}
    const request = {...envelope(document), payload: {mutations: [mutation]}}
    await client.apply(new Uint8Array([1]), document, request)
    expect(worker.requests.at(-1)).toMatchObject({payload: JSON.stringify(request.payload)})
    expect(() => client.apply(new Uint8Array([1]), document, {...request, payload: {mutations: [{...mutation, table: {rows: 2}}]}} as never)).toThrow()
    expect(() => client.apply(new Uint8Array([1]), document, {...request, payload: {mutations: [{...mutation, text: 'unsupported'}]}} as never)).toThrow()
  })

  it('validates hyperlink selections, URLs and wrapper anchors before posting', async () => {
    const worker = new FakeWorker(), client = createDocxWasmClient({workerFactory: () => worker})
    const document = structuredClone(fixtureDocument)
    const paragraph = document.body.blocks.find(block => block.paragraph)!.paragraph!
    if (!paragraph.edit_policy.allowed_operations.includes('hyperlink.set')) paragraph.edit_policy.allowed_operations.push('hyperlink.set')
    const run = paragraph.runs.find(run => run.kind === 'text')!
    run.can_edit_hyperlink = true
    const mutation = {target_kind: 'run' as const, target_id: run.id, expected_xml_sha256: run.anchor.xml_sha256, operation: 'hyperlink.set' as const, hyperlink: {url: 'https://example.com'}, range: {start_utf16: 0, end_utf16: 1}}
    const request = {...envelope(document), payload: {mutations: [mutation]}}
    await client.apply(new Uint8Array([1]), document, request)
    expect(worker.requests.at(-1)).toMatchObject({payload: JSON.stringify(request.payload)})
    for (const url of ['javascript:alert(1)', 'file:///tmp/example', 'https://user:password@example.com']) expect(() => client.apply(new Uint8Array([1]), document, {...request, payload: {mutations: [{...mutation, hyperlink: {url}}]}})).toThrow()
    expect(() => client.apply(new Uint8Array([1]), document, {...request, payload: {mutations: [{...mutation, hyperlink: {url: null, expected_xml_sha256: run.anchor.xml_sha256}}]}})).toThrow(/anchor/)
  })

  it('terminates on extraction JSON that fails the public Docs contract', async () => {
    const worker = new FakeWorker('{"version":1}')
    const client = createDocxWasmClient({ workerFactory: () => worker })
    await expect(client.extract(new Uint8Array([1]))).rejects.toBeInstanceOf(NativeDocxValidationError)
    expect(worker.terminated).toBe(true)
    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'TERMINATED', fatal: true })
  })

  it('throws synchronously on byte limits before creating or copying into a worker', () => {
    let calls = 0
    const packageClient = createDocxWasmClient({
      maxPackageBytes: 2,
      workerFactory: () => { calls++; return new FakeWorker() },
    })
    expect(() => packageClient.extract(new Uint8Array([1, 2, 3]))).toThrow(/maxPackageBytes is 2/)
    const payloadClient = createDocxWasmClient({ maxMutationPayloadBytes: 128, workerFactory: () => { calls++; return new FakeWorker() } })
    const large = envelope()
    large.payload.mutations[0].text = '🚀'.repeat(40)
    expect(() => payloadClient.apply(new Uint8Array([1]), fixtureDocument, large)).toThrow(/UTF-8 bytes/)
    expect(calls).toBe(0)
    expect(() => createDocxWasmClient({ maxPackageBytes: DOCX_WASM_NATIVE_MAX_PACKAGE_BYTES + 1 })).toThrow(/maxPackageBytes/)
  })

  it('strictly checks the public envelope, document CAS, and run anchor', () => {
    const client = createDocxWasmClient({ workerFactory: () => new FakeWorker() })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, { ...envelope(), expected_revision: 'sha256:' + '0'.repeat(64) })).toThrow(/expected_revision/)
    const stale = envelope()
    stale.payload.mutations[0].expected_xml_sha256 = 'sha256:' + '0'.repeat(64)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, stale)).toThrow(/run text anchor/)
    const unknown = { ...envelope(), extra: true } as unknown as NativeDocxOfficeMutationEnvelopeV1
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, unknown)).toThrow(/fields are invalid/)
    const accessor = envelope() as unknown as Record<string, unknown>
    Object.defineProperty(accessor, 'payload', { get: () => envelope().payload, enumerable: true })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, accessor as unknown as NativeDocxOfficeMutationEnvelopeV1)).toThrow(/accessor/)
    const invalidDocument = { ...fixtureDocument, protocol: 'wrong' } as unknown as NativeDocxDocumentV1
    expect(() => client.apply(new Uint8Array([1]), invalidDocument, envelope())).toThrow(NativeDocxValidationError)
  })

  it('accepts the Go-proven single-text-run paragraph target with its paragraph anchor', async () => {
    const worker = new FakeWorker()
    const client = createDocxWasmClient({ workerFactory: () => worker })
    const paragraphEnvelope = envelope()
    paragraphEnvelope.payload.mutations[0] = {
      target_kind: 'paragraph',
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      text: 'Strict paragraph',
    }

    await expect(client.apply(new Uint8Array([1]), fixtureDocument, paragraphEnvelope)).resolves.toEqual(new Uint8Array([4, 5, 6]))
    expect(worker.requests[1]).toMatchObject({ payload: JSON.stringify(paragraphEnvelope.payload) })
  })

  it('carries a run-property patch, with its range, to the engine', async () => {
    const worker = new FakeWorker()
    const client = createDocxWasmClient({ workerFactory: () => worker })
    const formatting = { ...envelope(), payload: { mutations: [{
      target_kind: 'run' as const,
      target_id: fixtureRun.id,
      expected_xml_sha256: fixtureRun.anchor.xml_sha256,
      properties: { bold: true, color: 'ff0000', font_size_half_points: 24 },
      range: { start_utf16: 0, end_utf16: 3 },
    }] } }
    await expect(client.apply(new Uint8Array([1]), fixtureDocument, formatting)).resolves.toEqual(new Uint8Array([4, 5, 6]))
    // The colour is normalised to the upper-case form the contract reads back.
    expect(JSON.parse((worker.requests[1] as { payload: string }).payload)).toEqual({ mutations: [{
      target_kind: 'run',
      target_id: fixtureRun.id,
      expected_xml_sha256: fixtureRun.anchor.xml_sha256,
      properties: { bold: true, color: 'FF0000', font_size_half_points: 24 },
      range: { start_utf16: 0, end_utf16: 3 },
    }] })
  })

  it('formats a paragraph target by its own anchor, multiple runs included', async () => {
    const worker = new FakeWorker()
    const client = createDocxWasmClient({ workerFactory: () => worker })
    const formatting = { ...envelope(multiRunDocument), payload: { mutations: [{
      target_kind: 'paragraph' as const,
      target_id: multiRunParagraph.id,
      expected_xml_sha256: multiRunParagraph.anchor.xml_sha256,
      properties: { italic: true },
    }] } }
    await expect(client.apply(new Uint8Array([1]), multiRunDocument, formatting)).resolves.toEqual(new Uint8Array([4, 5, 6]))
  })

  it('refuses formatting that the engine would refuse or that mixes shapes', () => {
    const client = createDocxWasmClient({ workerFactory: () => new FakeWorker() })
    const format = (properties: unknown, range?: unknown) => ({ ...envelope(), payload: { mutations: [{
      target_kind: 'run' as const,
      target_id: fixtureRun.id,
      expected_xml_sha256: fixtureRun.anchor.xml_sha256,
      properties,
      ...(range === undefined ? {} : { range }),
    }] } } as unknown as NativeDocxOfficeMutationEnvelopeV1)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({}))).toThrow(/sets no run property/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ bold: 'yes' }))).toThrow(/bold must be a boolean/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ underline: 'squiggly' }))).toThrow(/underline is outside/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ highlight: 'chartreuse' }))).toThrow(/highlight is outside/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ color: '#ff0000' }))).toThrow(/six hex digits/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ font_size_half_points: 1 }))).toThrow(/2\.\.3276/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ font_family: 'Bad<Font>' }))).toThrow(/bounded font name/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ bold: true }, { start_utf16: 2, end_utf16: 2 }))).toThrow(/non-empty UTF-16 span/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, format({ bold: true, weight: 700 }))).toThrow(/unknown key/)
    const stale = format({ bold: true }) as unknown as { payload: { mutations: Array<{ expected_xml_sha256: string }> } }
    stale.payload.mutations[0].expected_xml_sha256 = 'sha256:' + '0'.repeat(64)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, stale as unknown as NativeDocxOfficeMutationEnvelopeV1)).toThrow(/does not match an extracted run anchor/)
    const mixed = { ...envelope(), payload: { mutations: [
      { target_kind: 'run' as const, target_id: fixtureRun.id, expected_xml_sha256: fixtureRun.anchor.xml_sha256, text: 'Replaced' },
      { target_kind: 'run' as const, target_id: fixtureRun.id, expected_xml_sha256: fixtureRun.anchor.xml_sha256, properties: { bold: true } },
    ] } } as unknown as NativeDocxOfficeMutationEnvelopeV1
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, mixed)).toThrow(/may not mix text replacements/)
  })

  it('carries paragraph alignment on its own paragraph anchor', async () => {
    const worker = new FakeWorker()
    const client = createDocxWasmClient({ workerFactory: () => worker })
    const aligned = { ...envelope(), payload: { mutations: [{
      target_kind: 'paragraph' as const,
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      properties: { alignment: 'center' as const },
    }] } }
    await expect(client.apply(new Uint8Array([1]), fixtureDocument, aligned)).resolves.toEqual(new Uint8Array([4, 5, 6]))
    const alignment = (properties: unknown, extra: Record<string, unknown> = {}) => ({ ...envelope(), payload: { mutations: [{
      target_kind: 'paragraph' as const,
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      properties,
      ...extra,
    }] } } as unknown as NativeDocxOfficeMutationEnvelopeV1)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, alignment({ alignment: 'middle' }))).toThrow(/Invalid paragraph alignment/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, alignment({ alignment: 'center', bold: true }))).toThrow(/unknown|unsupported/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, alignment({ alignment: 'center' }, { range: { start_utf16: 0, end_utf16: 2 } }))).toThrow(/whole paragraph target/)
    const layout = { spacing_before_twips: 120, indent_left_twips: -120, first_line_twips: null, hanging_twips: 360, line_spacing: 480, line_rule: 'auto' }
    await expect(client.apply(new Uint8Array([1]), fixtureDocument, alignment(layout))).resolves.toEqual(new Uint8Array([4, 5, 6]))
    for (const invalid of [{ line_spacing: 1.5 }, { indent_left_twips: -31681 }, { line_rule: 'bad' }, { first_line_twips: 1, hanging_twips: 1 }]) {
      expect(() => client.apply(new Uint8Array([1]), fixtureDocument, alignment(invalid))).toThrow()
    }
    const onRun = { ...envelope(), payload: { mutations: [{
      target_kind: 'run' as const,
      target_id: fixtureRun.id,
      expected_xml_sha256: fixtureRun.anchor.xml_sha256,
      properties: { alignment: 'center' },
    }] } } as unknown as NativeDocxOfficeMutationEnvelopeV1
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, onRun)).toThrow(/whole paragraph target/)
  })

  it('carries paragraph numbering on its own paragraph anchor', async () => {
    const worker = new FakeWorker()
    const client = createDocxWasmClient({ workerFactory: () => worker })
    const numbered = { ...envelope(), payload: { mutations: [{
      target_kind: 'paragraph' as const,
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      properties: { numbering_kind: 'bullet' as const },
    }] } }
    await expect(client.apply(new Uint8Array([1]), fixtureDocument, numbered)).resolves.toEqual(new Uint8Array([4, 5, 6]))
    expect(JSON.parse((worker.requests[1] as { payload: string }).payload)).toEqual({ mutations: [{
      target_kind: 'paragraph',
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      properties: { numbering_kind: 'bullet' },
    }] })
    const numbering = (properties: unknown, extra: Record<string, unknown> = {}) => ({ ...envelope(), payload: { mutations: [{
      target_kind: 'paragraph' as const,
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      properties,
      ...extra,
    }] } } as unknown as NativeDocxOfficeMutationEnvelopeV1)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, numbering({ numbering_kind: 'outline' }))).toThrow(/numbering_kind must be bullet or decimal/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, numbering({ numbering_kind: 'bullet', bold: true }))).toThrow(/numbering on its own/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, numbering({ numbering_kind: 'bullet' }, { range: { start_utf16: 0, end_utf16: 2 } }))).toThrow(/takes no range/)
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, numbering({ numbering_num_id: '1', numbering_level: 9 }))).toThrow(/0\.\.8/)
    const onRun = { ...envelope(), payload: { mutations: [{
      target_kind: 'run' as const,
      target_id: fixtureRun.id,
      expected_xml_sha256: fixtureRun.anchor.xml_sha256,
      properties: { numbering_kind: 'bullet' },
    }] } } as unknown as NativeDocxOfficeMutationEnvelopeV1
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, onRun)).toThrow(/needs a paragraph target/)
  })

  it('refuses paragraph targets with multiple runs and paragraph/run overlap', () => {
    const client = createDocxWasmClient({ workerFactory: () => new FakeWorker() })
    const unsupported = envelope(multiRunDocument)
    unsupported.payload.mutations[0] = {
      target_kind: 'paragraph',
      target_id: multiRunParagraph.id,
      expected_xml_sha256: multiRunParagraph.anchor.xml_sha256,
      text: 'Flattened',
    }
    expect(() => client.apply(new Uint8Array([1]), multiRunDocument, unsupported)).toThrow(/paragraph text anchor/)

    const overlap = envelope()
    overlap.payload.mutations.unshift({
      target_kind: 'paragraph',
      target_id: fixtureParagraph.id,
      expected_xml_sha256: fixtureParagraph.anchor.xml_sha256,
      text: 'Paragraph text',
    })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, overlap)).toThrow(/overlaps another paragraph or run target/)
  })

  it('refuses duplicate targets and invalid XML text', () => {
    const client = createDocxWasmClient({ workerFactory: () => new FakeWorker() })
    const duplicate = envelope()
    duplicate.payload.mutations.push({ ...duplicate.payload.mutations[0] })
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, duplicate)).toThrow(/same native target/)
    const control = envelope()
    control.payload.mutations[0].text = 'bad\u0000text'
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, control)).toThrow(/XML 1.0/)
    const carriageReturn = envelope()
    carriageReturn.payload.mutations[0].text = 'normalized\rtext'
    expect(() => client.apply(new Uint8Array([1]), fixtureDocument, carriageReturn)).toThrow(/XML 1.0/)
  })

  it('keeps module construction SSR-safe and reports missing Worker on use', async () => {
    const client = createDocxWasmClient()
    await expect(client.extract(new Uint8Array([1]))).rejects.toMatchObject({ code: 'WORKER_UNAVAILABLE', fatal: true })
  })
})
