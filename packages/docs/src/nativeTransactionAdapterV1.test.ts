import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { NativeDocxDocumentV1 } from './nativeContract.js'
import {
  DOCX_PROSEMIRROR_PROJECTION_SCHEMA,
  DOCX_PROSEMIRROR_PROJECTION_SCHEMA_VERSION,
  DOCX_PROSEMIRROR_TRANSACTION_PROTOCOL,
  DOCX_PROSEMIRROR_TRANSACTION_VERSION,
  DOCX_TRANSACTION_ADAPTER_LIMITS,
  adaptNativeDocxProseMirrorTransactionV1,
  adaptNativeDocxProseMirrorTransactionWithHostV1,
  decodeNativeDocxProseMirrorTransactionJsonV1,
  type NativeDocxProseMirrorDocumentProjectionV1,
  type NativeDocxProseMirrorReplaceStepV1,
  type NativeDocxProseMirrorTransactionV1,
  type NativeDocxTextMutationPayloadV1,
} from './nativeTransactionAdapterV1.js'

const expectedRecord = JSON.parse(readFileSync(new URL('../../../go/officecompat/corpus/generated/expected/docx-strict-relocated.json', import.meta.url), 'utf8')) as { native: NativeDocxDocumentV1 }
const exactVector = JSON.parse(readFileSync(new URL('../../../testdata/docx-native/transaction-adapter-envelope-v1.json', import.meta.url), 'utf8')) as { expected_text: string; encoded_envelope: string }
const baseDocument = expectedRecord.native
const marks = ['font-size:24', 'italic']

function projection(document: NativeDocxDocumentV1, runMarks = marks): NativeDocxProseMirrorDocumentProjectionV1 {
  let cursor = 0
  const paragraphs = document.body.blocks.map((block) => {
      const paragraph = block.paragraph!
      const nodeFrom = cursor
      let textCursor = nodeFrom + 1
      const runs = paragraph.runs.map((run) => {
        const text = run.text!
        const result = {
          run_id: run.id,
          text_from: textCursor,
          text_to: textCursor + text.length,
          text,
          marks: [...runMarks],
          native_properties: { ...(run.properties ?? {}) },
        }
        textCursor = result.text_to
        return result
      })
      cursor = textCursor + 1
      return { paragraph_id: paragraph.id, node_from: nodeFrom, node_to: cursor, runs }
    })
  return { doc_size: cursor, paragraphs }
}

function replaceInProjection(before: NativeDocxProseMirrorDocumentProjectionV1, runId: string, from: number, to: number, inserted: string): NativeDocxProseMirrorDocumentProjectionV1 {
  const after = structuredClone(before)
  const flat = after.paragraphs.flatMap((paragraph) => paragraph.runs)
  const runIndex = flat.findIndex((run) => run.run_id === runId)
  const run = flat[runIndex]!
  const oldLength = run.text.length
  const localFrom = from - run.text_from
  const localTo = to - run.text_from
  run.text = run.text.slice(0, localFrom) + inserted + run.text.slice(localTo)
  const delta = run.text.length - oldLength
  run.text_to += delta
  let passed = false
  for (const paragraph of after.paragraphs) {
    const contains = paragraph.runs.some((entry) => entry.run_id === runId)
    for (const entry of paragraph.runs) {
      if (passed && entry.run_id !== runId) {
        entry.text_from += delta
        entry.text_to += delta
      }
    }
    if (contains) {
      paragraph.node_to += delta
      passed = true
    } else if (passed) {
      paragraph.node_from += delta
      paragraph.node_to += delta
    }
  }
  after.doc_size += delta
  return after
}

function step(before: NativeDocxProseMirrorDocumentProjectionV1, index: number, from: number, to: number, inserted: string): { step: NativeDocxProseMirrorReplaceStepV1; after: NativeDocxProseMirrorDocumentProjectionV1 } {
  const paragraph = before.paragraphs.find((entry) => entry.runs.some((run) => from >= run.text_from && to <= run.text_to))!
  const run = paragraph.runs.find((entry) => from >= entry.text_from && to <= entry.text_to)!
  const deleted = run.text.slice(from - run.text_from, to - run.text_from)
  const after = replaceInProjection(before, run.run_id, from, to, inserted)
  const afterRun = after.paragraphs.flatMap((entry) => entry.runs).find((entry) => entry.run_id === run.run_id)!
  return {
    step: {
      step_type: 'replace', step_index: index, from, to,
      paragraph_id: paragraph.paragraph_id, run_id: run.run_id,
      expected_xml_sha256: baseDocument.body.blocks[0]!.paragraph!.runs[0]!.anchor.xml_sha256,
      deleted_text: deleted, before_text: run.text, after_text: afterRun.text,
      slice: { open_start: 0, open_end: 0, text: inserted, marks: inserted === '' ? [] : [...run.marks] },
    },
    after,
  }
}

function transaction(before: NativeDocxProseMirrorDocumentProjectionV1, steps: NativeDocxProseMirrorReplaceStepV1[], after: NativeDocxProseMirrorDocumentProjectionV1, document = baseDocument): NativeDocxProseMirrorTransactionV1 {
  return {
    protocol: DOCX_PROSEMIRROR_TRANSACTION_PROTOCOL,
    version: DOCX_PROSEMIRROR_TRANSACTION_VERSION,
    schema_id: DOCX_PROSEMIRROR_PROJECTION_SCHEMA,
    schema_version: DOCX_PROSEMIRROR_PROJECTION_SCHEMA_VERSION,
    mutation_id: 'pm-save-1',
    document_id: document.document_id,
    expected_revision: document.source.package_sha256,
    before,
    steps,
    after,
  }
}

function oneStep(document: NativeDocxDocumentV1, from: number, to: number, inserted: string): NativeDocxProseMirrorTransactionV1 {
  const before = projection(document)
  const projected = step(before, 0, from, to, inserted)
  return transaction(before, [projected.step], projected.after, document)
}

function expectRefusal(result: ReturnType<typeof adaptNativeDocxProseMirrorTransactionV1>, code: string): void {
  expect(result.ok).toBe(false)
  if (result.ok) return
  expect(result.issues).toEqual(expect.arrayContaining([expect.objectContaining({ code })]))
  expect(result).not.toHaveProperty('value')
  expect(result).not.toHaveProperty('envelope')
}

describe('native DOCX ProseMirror transaction adapter v1', () => {
  it('types the native writer single-text-run paragraph selector', () => {
    const paragraph = baseDocument.body.blocks[0]!.paragraph!
    const payload: NativeDocxTextMutationPayloadV1 = { mutations: [{
      target_kind: 'paragraph',
      target_id: paragraph.id,
      expected_xml_sha256: paragraph.anchor.xml_sha256,
      text: 'Strict paragraph',
    }] }
    expect(payload.mutations[0]).toMatchObject({ target_kind: 'paragraph', target_id: paragraph.id })
  })

  it.each([
    ['insert', 7, 7, 'ly', 'Strictly routed'],
    ['delete', 7, 14, '', 'Strict'],
    ['replace', 8, 14, 'native', 'Strict native'],
  ])('maps a provable UTF-16 %s to one guarded native run replacement', (_name, from, to, inserted, expected) => {
    const result = adaptNativeDocxProseMirrorTransactionV1(baseDocument, oneStep(baseDocument, from, to, inserted))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.envelope.payload.mutations).toEqual([{
      target_kind: 'run',
      target_id: 'run:eb2d52f19c42f8817e1d56fa',
      expected_xml_sha256: 'sha256:0e0d3ceb49fbb4e3159a96668c911c0ebf42ca5a0f4f5113b892354bafa170c1',
      text: expected,
    }])
  })

  it('uses exact UTF-16 semantics without normalizing surrogate or combining sequences', () => {
    const document = structuredClone(baseDocument)
    document.body.blocks[0]!.paragraph!.runs[0]!.text = 'A😀e\u0301Z'
    const combining = adaptNativeDocxProseMirrorTransactionV1(document, oneStep(document, 6, 6, '́'))
    expect(combining.ok).toBe(true)
    if (combining.ok) expect(combining.value.envelope.payload.mutations[0]!.text).toBe('A😀é́Z')

    const emoji = adaptNativeDocxProseMirrorTransactionV1(document, oneStep(document, 2, 4, '🙂'))
    expect(emoji.ok).toBe(true)
    if (emoji.ok) expect(emoji.value.envelope.payload.mutations[0]!.text).toBe('A🙂éZ')

    const split = oneStep(document, 3, 3, 'x')
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, split), 'INVALID_VALUE')
    const lone = oneStep(document, 5, 6, '\ud800')
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, lone), 'INVALID_VALUE')
    for (const forbidden of ['\ufffe', '\uffff']) {
      expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, oneStep(document, 4, 4, forbidden)), 'INVALID_VALUE')
    }
  })

  it('validates multiple successive coordinates and coalesces ordered steps per run', () => {
    const document = structuredClone(baseDocument)
    document.body.blocks[0]!.paragraph!.runs[0]!.text = 'abcdef'
    const before = projection(document)
    const first = step(before, 0, 2, 2, 'X')
    const second = step(first.after, 1, 6, 7, 'E')
    const input = transaction(before, [first.step, second.step], second.after, document)
    const result = adaptNativeDocxProseMirrorTransactionV1(document, input)
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value.envelope.payload.mutations).toEqual([expect.objectContaining({ text: 'aXbcdEf' })])
  })

  it('shifts later paragraphs and emits multi-target mutations in native order', () => {
    const document = structuredClone(baseDocument)
    const secondBlock = structuredClone(document.body.blocks[0]!)
    secondBlock.id = 'paragraph:second'
    secondBlock.paragraph!.id = 'paragraph:second'
    secondBlock.paragraph!.anchor.path = '/w:document[1]/w:body[1]/w:p[2]'
    secondBlock.paragraph!.anchor.xml_sha256 = `sha256:${'1'.repeat(64)}`
    secondBlock.paragraph!.runs[0]!.id = 'run:second'
    secondBlock.paragraph!.runs[0]!.anchor.path = '/w:document[1]/w:body[1]/w:p[2]/w:r[1]/w:t[1]'
    secondBlock.paragraph!.runs[0]!.anchor.xml_sha256 = `sha256:${'2'.repeat(64)}`
    secondBlock.paragraph!.runs[0]!.text = 'Second'
    document.body.blocks.push(secondBlock)

    const before = projection(document)
    const first = step(before, 0, 8, 14, 'native')
    const secondRun = first.after.paragraphs[1]!.runs[0]!
    const second = step(first.after, 1, secondRun.text_from + 6, secondRun.text_from + 6, '!')
    second.step.expected_xml_sha256 = secondBlock.paragraph!.runs[0]!.anchor.xml_sha256
    const result = adaptNativeDocxProseMirrorTransactionV1(document, transaction(before, [first.step, second.step], second.after, document))
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.value.envelope.payload.mutations).toEqual([
      expect.objectContaining({ target_id: baseDocument.body.blocks[0]!.paragraph!.runs[0]!.id, text: 'Strict native' }),
      expect.objectContaining({ target_id: 'run:second', text: 'Second!' }),
    ])
  })

  it('rejects mapping drift, overlap, revisited insertion points, and reordered steps atomically', () => {
    const document = structuredClone(baseDocument)
    document.body.blocks[0]!.paragraph!.runs[0]!.text = 'abcdef'
    const before = projection(document)
    const drift = step(before, 0, 2, 3, 'B')
    drift.step.before_text = 'stale'
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, transaction(before, [drift.step], drift.after, document)), 'MAPPING_DRIFT')

    const first = step(before, 0, 2, 4, 'BC')
    const overlap = step(first.after, 1, 3, 5, 'CD')
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, transaction(before, [first.step, overlap.step], overlap.after, document)), 'OVERLAPPING_STEPS')

    const insertion = step(before, 0, 2, 2, 'X')
    const revisit = step(insertion.after, 1, 2, 2, 'Y')
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, transaction(before, [insertion.step, revisit.step], revisit.after, document)), 'OVERLAPPING_STEPS')

    const later = step(before, 0, 5, 6, 'E')
    const earlier = step(later.after, 1, 2, 3, 'B')
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, transaction(before, [later.step, earlier.step], earlier.after, document)), 'REORDERED_STEPS')
  })

  it('rejects stale revisions, document/run IDs, anchors, schema drift, marks, and partial coverage', () => {
    const cases: Array<[string, (input: any) => void, string]> = [
      ['revision', (input) => { input.expected_revision = `sha256:${'0'.repeat(64)}` }, 'STALE_REVISION'],
      ['document id', (input) => { input.document_id = 'document:stale' }, 'STALE_TARGET'],
      ['run id', (input) => { input.steps[0].run_id = 'run:stale' }, 'STALE_TARGET'],
      ['anchor', (input) => { input.steps[0].expected_xml_sha256 = `sha256:${'0'.repeat(64)}` }, 'STALE_TARGET'],
      ['schema', (input) => { input.schema_version = 2 }, 'SCHEMA_DRIFT'],
      ['structural step', (input) => { input.steps[0].step_type = 'replaceAround' }, 'UNSUPPORTED_STRUCTURE'],
      ['marks', (input) => { input.steps[0].slice.marks = ['bold'] }, 'AMBIGUOUS_EDIT'],
      ['after marks', (input) => { input.after.paragraphs[0].runs[0].marks = ['bold'] }, 'PARTIAL_COVERAGE'],
      ['after properties', (input) => { input.after.paragraphs[0].runs[0].native_properties.bold = input.after.paragraphs[0].runs[0].native_properties.bold !== true }, 'PARTIAL_COVERAGE'],
      ['coverage', (input) => { input.after.paragraphs[0].runs[0].text = 'fabricated' }, 'PARTIAL_COVERAGE'],
    ]
    for (const [_name, mutate, code] of cases) {
      const input: any = oneStep(baseDocument, 8, 14, 'native')
      mutate(input)
      expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, input), code)
    }
  })

  it('rejects structural documents, tables, drawings, notes, fields, tracked changes, and unsafe whitespace', () => {
    const tableRecord = JSON.parse(readFileSync(new URL('../../../go/officecompat/corpus/generated/expected/docx-transitional-common.json', import.meta.url), 'utf8')) as { native: NativeDocxDocumentV1 }
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(tableRecord.native, {}), 'UNSUPPORTED_STRUCTURE')

    for (const mutate of [
      (document: any) => {
        const note = structuredClone(document.headers[0])
        note.id = 'story:note'; note.kind = 'footnote'; note.native_story_id = '-1'; note.relationship_id = 'rIdFootnotes'; note.note_role = 'separator'; note.part_name = 'word/footnotes.xml'; note.anchor.part_name = note.part_name
        note.blocks[0].id = 'paragraph:note'; note.blocks[0].paragraph.id = 'paragraph:note'; note.blocks[0].paragraph.anchor.part_name = note.part_name
        note.blocks[0].paragraph.runs[0].id = 'run:note'; note.blocks[0].paragraph.runs[0].anchor.part_name = note.part_name
        document.notes = [note]
      },
      (document: any) => { document.unsupported = [{ id: 'unsupported:x', code: 'FIELD_SEMANTICS', capability: 'fields', scope_id: document.body.id, preservation: 'refuse-mutation', message: 'field' }] },
      (document: any) => { document.unsupported = [{ id: 'unsupported:x', code: 'WRAPPED_RUN_MARKUP', capability: 'tracked-changes', scope_id: document.body.id, preservation: 'refuse-mutation', message: 'tracked' }] },
      (document: any) => { document.headers[0].blocks[0].paragraph.runs[0].kind = 'drawing'; document.headers[0].blocks[0].paragraph.runs[0].text = undefined; document.headers[0].blocks[0].paragraph.runs[0].drawing = { id: 'drawing:x', anchor: document.headers[0].blocks[0].paragraph.runs[0].anchor, placement: 'inline', width_emu: 1, height_emu: 1, edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'NO', message: 'no', preservation: 'refuse-mutation' } } } },
      (document: any) => { const source = document.body.blocks[0].paragraph.runs[0]; document.body.blocks[0].paragraph.runs.push({ kind: 'control', id: 'run:control', anchor: structuredClone(source.anchor), properties: {}, control: 'tab' }) },
    ]) {
      const document: any = structuredClone(baseDocument)
      mutate(document)
      expectRefusal(adaptNativeDocxProseMirrorTransactionV1(document, oneStep(baseDocument, 8, 14, 'native')), 'UNSUPPORTED_STRUCTURE')
    }
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, oneStep(baseDocument, 1, 1, ' ')), 'AMBIGUOUS_EDIT')
    const preservedWhitespace = structuredClone(baseDocument)
    preservedWhitespace.body.blocks[0]!.paragraph!.runs[0]!.text = ' Strict routed '
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(preservedWhitespace, oneStep(preservedWhitespace, 9, 15, 'native')), 'AMBIGUOUS_EDIT')
  })

  it('rejects hostile objects, unknown keys, invalid JSON scalars, XML controls, and resource excess with no output', () => {
    const unknown: any = oneStep(baseDocument, 8, 14, 'native')
    unknown.steps[0].host_secret = true
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, unknown), 'UNKNOWN_FIELD')

    const accessor: any = oneStep(baseDocument, 8, 14, 'native')
    Object.defineProperty(accessor.steps[0], 'from', { get() { throw new Error('hostile') }, enumerable: true })
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, accessor), 'INVALID_INPUT')

    const cyclic: any = oneStep(baseDocument, 8, 14, 'native')
    cyclic.after.loop = cyclic
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, cyclic), 'LIMIT_EXCEEDED')

    const xmlControl = oneStep(baseDocument, 8, 14, '\u0001')
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, xmlControl), 'INVALID_VALUE')

    const excessive: any = oneStep(baseDocument, 8, 14, 'native')
    excessive.steps = Array.from({ length: DOCX_TRANSACTION_ADAPTER_LIMITS.maxSteps + 1 }, () => excessive.steps[0])
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, excessive), 'LIMIT_EXCEEDED')
    const hugeText = oneStep(baseDocument, 8, 14, 'x'.repeat(1_048_577))
    expectRefusal(adaptNativeDocxProseMirrorTransactionV1(baseDocument, hugeText), 'LIMIT_EXCEEDED')
    expect(decodeNativeDocxProseMirrorTransactionJsonV1(' '.repeat(DOCX_TRANSACTION_ADAPTER_LIMITS.maxJsonBytes + 1))).toEqual([expect.objectContaining({ code: 'LIMIT_EXCEEDED' })])
  })

  it('is deterministic, emits the shared exact canonical envelope bytes, and exposes only an optional host projection seam', () => {
    const input = oneStep(baseDocument, 8, 14, 'native')
    const first = adaptNativeDocxProseMirrorTransactionV1(structuredClone(baseDocument), structuredClone(input))
    const second = adaptNativeDocxProseMirrorTransactionV1(structuredClone(baseDocument), structuredClone(input))
    expect(first).toEqual(second)
    expect(first.ok).toBe(true)
    if (!first.ok) return
    expect(first.value.encoded_envelope).toBe(exactVector.encoded_envelope)
    expect(new TextEncoder().encode(first.value.encoded_envelope).byteLength).toBeLessThanOrEqual(DOCX_TRANSACTION_ADAPTER_LIMITS.maxJsonBytes)
    const host = adaptNativeDocxProseMirrorTransactionWithHostV1(baseDocument, { transaction: 'opaque' }, { project: () => input })
    expect(host).toEqual(first)
  })

  it('isolates the authoritative document from host projector mutation', () => {
    const document = structuredClone(baseDocument)
    const input = oneStep(document, 8, 14, 'native')
    const before = structuredClone(document)
    const result = adaptNativeDocxProseMirrorTransactionWithHostV1(document, {}, {
      project: (_transaction, projectorDocument) => {
        projectorDocument.source.package_sha256 = `sha256:${'0'.repeat(64)}`
        projectorDocument.body.blocks[0]!.paragraph!.runs[0]!.text = 'forged'
        return input
      },
    })
    expect(result.ok).toBe(true)
    expect(document).toEqual(before)
  })

  it('rejects duplicate JSON members at every transaction object layer', () => {
    const encoded = JSON.stringify(oneStep(baseDocument, 8, 14, 'native'))
    const cases = [
      encoded.replace('"version":1', '"version":1,"version":1'),
      encoded.replace('"step_index":0', '"step_index":0,"step_index":0'),
      encoded.replace('"open_start":0', '"open_start":0,"open_start":0'),
      encoded.replace('"doc_size":15', '"doc_size":15,"doc_size":15'),
    ]
    for (const json of cases) {
      const decoded = decodeNativeDocxProseMirrorTransactionJsonV1(json)
      expect(Array.isArray(decoded)).toBe(true)
      if (Array.isArray(decoded)) expect(decoded).toEqual([expect.objectContaining({ code: 'UNKNOWN_FIELD' })])
    }
  })

  it('contains no structural, HTML, DOM, Mammoth, or reconstruction authority', () => {
    const source = readFileSync(new URL('./nativeTransactionAdapterV1.ts', import.meta.url), 'utf8').toLowerCase()
    expect(source).not.toMatch(/(?:mammoth|domparser|innerhtml|document\.queryselector|archive\/zip|encoding\/xml)/)
    expect(source).not.toMatch(/from\s+['"](?:@tiptap|prosemirror|react|jszip)/)
  })
})
