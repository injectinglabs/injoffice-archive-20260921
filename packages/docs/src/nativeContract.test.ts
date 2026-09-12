import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  DOCX_MAX_TWIPS_FOR_MILLIPOINTS,
  DOCX_NATIVE_PROTOCOL,
  DOCX_NATIVE_LIMITS,
  DOCX_NATIVE_V1_BINDING_FIELDS,
  NativeDocxValidationError,
  decodeNativeDocxDocument,
  decodeNativeDocxJson,
  encodeNativeDocxDocument,
  type NativeDocxDocumentV1,
} from './nativeContract.js'

const fixture = JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json', import.meta.url), 'utf8')) as NativeDocxDocumentV1
const invalidFixture = JSON.parse(readFileSync(new URL('../../../testdata/docx-native/invalid-v1.json', import.meta.url), 'utf8')) as {
  cases: Array<{ name: string; pointer: string; value: unknown; code: string; path: string }>
}

function setPointer(root: Record<string, any>, pointer: string, value: unknown): void {
  const segments = pointer.slice(1).split('/').map((segment) => segment.replace(/~1/g, '/').replace(/~0/g, '~'))
  const key = segments.pop()
  let owner: any = root
  for (const segment of segments) owner = owner[segment]
  owner[key!] = value
}

describe('native DOCX contract v1', () => {
  it('bounds inline effect extents and rejects malformed or floating projections', () => {
    for (const effect of [{left:0,top:0,right:0}, {left:-1,top:0,right:0,bottom:0}, {left:0.5,top:0,right:0,bottom:0}, {left:91_440_001,top:0,right:0,bottom:0}, {left:0,top:0,right:0,bottom:0,extra:1}]) {
      const value = structuredClone(fixture)
      const drawing = value.body.blocks[0]!.paragraph!.runs.find(run => run.drawing)!.drawing!
      drawing.inline_effect_extent_emu = effect as any
      expect(decodeNativeDocxDocument(value).ok).toBe(false)
    }
    const value = structuredClone(fixture)
    const drawing = value.body.blocks[0]!.paragraph!.runs.find(run => run.drawing)!.drawing!
    drawing.inline_effect_extent_emu = {left:0,top:0,right:0,bottom:0}
    drawing.placement = 'floating'
    expect(decodeNativeDocxDocument(value).ok).toBe(false)
  })
  it('accepts the cross-language fixture with native stories, references, tables, and drawings', () => {
    const decoded = decodeNativeDocxDocument(fixture)
    expect(decoded.ok).toBe(true)
    if (!decoded.ok) return
    expect(decoded.value.protocol).toBe(DOCX_NATIVE_PROTOCOL)
    expect(decoded.value.body.blocks.map((block) => block.kind)).toEqual(['paragraph', 'table'])
    expect(decoded.value.sections[0]?.header_refs[0]?.story_id).toBe('story:header:default')
    expect(decoded.value.comment_stories[0]?.blocks[0]?.kind).toBe('paragraph')
  })

  it('classifies OPC media MIME types with ASCII case-insensitive semantics', () => {
    const value = structuredClone(fixture)
    value.passthrough_parts.find((part) => part.part_name === 'word/media/image1.png')!.content_type = 'IMAGE/PNG'
    expect(decodeNativeDocxDocument(value)).toEqual(expect.objectContaining({ ok: true }))
  })

  it('serializes object keys deterministically without changing source-order arrays', () => {
    const value = structuredClone(fixture)
    const scrambled = {
      unsupported: value.unsupported,
      passthrough_parts: value.passthrough_parts,
      capabilities: value.capabilities,
      comments: value.comments,
      comment_stories: value.comment_stories,
      notes: value.notes,
      footers: value.footers,
      headers: value.headers,
      sections: value.sections,
      body: value.body,
      source: value.source,
      revision: value.revision,
      document_id: value.document_id,
      version: value.version,
      protocol: value.protocol,
    }
    expect(encodeNativeDocxDocument(scrambled)).toBe(encodeNativeDocxDocument(value))
    expect(JSON.parse(encodeNativeDocxDocument(value)).body.blocks.map((block: { id: string }) => block.id)).toEqual([
      'paragraph:intro',
      'table:summary',
    ])
  })

  it('binds stable column identity and geometry into canonical encoding', () => {
    const changed = structuredClone(fixture)
    changed.sections[0]!.page.column_definitions[0]!.id = 'column:section:1:changed'
    expect(encodeNativeDocxDocument(changed)).not.toBe(encodeNativeDocxDocument(fixture))

    const duplicate = structuredClone(fixture)
    duplicate.sections[0]!.page.columns = 2
    duplicate.sections[0]!.page.column_definitions.push({ id: duplicate.sections[0]!.page.column_definitions[0]!.id, ordinal: 1 })
    expect(decodeNativeDocxDocument(duplicate)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ code: 'DUPLICATE_ID', path: '/sections/0/page/column_definitions/1/id' })]),
    }))

    const malformed = structuredClone(fixture) as any
    malformed.sections[0].page.column_layout = 'explicit'
    expect(decodeNativeDocxDocument(malformed)).toEqual(expect.objectContaining({ ok: false }))

    const unbounded = structuredClone(fixture) as any
    unbounded.sections[0].page.columns = 46
    expect(decodeNativeDocxDocument(unbounded)).toEqual(expect.objectContaining({
      ok: false,
      issues: expect.arrayContaining([expect.objectContaining({ code: 'OUT_OF_RANGE', path: '/sections/0/page/columns' })]),
    }))
  })

  it('publishes and enforces the exact x50 twip boundary across page and table geometry', () => {
    const exact = structuredClone(fixture)
    const table = exact.body.blocks[1]!.table!
    table.width_twips = DOCX_MAX_TWIPS_FOR_MILLIPOINTS
    table.grid_widths_twips = [DOCX_MAX_TWIPS_FOR_MILLIPOINTS]
    table.cell_margins = { top_twips: 0, right_twips: 0, bottom_twips: 0, left_twips: DOCX_MAX_TWIPS_FOR_MILLIPOINTS }
    exact.sections[0]!.page.width_twips = DOCX_MAX_TWIPS_FOR_MILLIPOINTS
    expect(decodeNativeDocxDocument(exact)).toEqual(expect.objectContaining({ ok: true }))

    for (const mutate of [
      (value: NativeDocxDocumentV1) => { value.body.blocks[1]!.table!.width_twips = DOCX_MAX_TWIPS_FOR_MILLIPOINTS + 1 },
      (value: NativeDocxDocumentV1) => { value.body.blocks[1]!.table!.grid_widths_twips = [DOCX_MAX_TWIPS_FOR_MILLIPOINTS + 1] },
      (value: NativeDocxDocumentV1) => { value.body.blocks[1]!.table!.cell_margins = { top_twips: 0, right_twips: 0, bottom_twips: 0, left_twips: DOCX_MAX_TWIPS_FOR_MILLIPOINTS + 1 } },
      (value: NativeDocxDocumentV1) => { value.sections[0]!.page.width_twips = DOCX_MAX_TWIPS_FOR_MILLIPOINTS + 1 },
    ]) {
      const unsafe = structuredClone(fixture)
      mutate(unsafe)
      expect(decodeNativeDocxDocument(unsafe)).toEqual(expect.objectContaining({ ok: false }))
    }
  })

  it('rejects unknown fields at modeled nesting levels', () => {
    const value = structuredClone(fixture) as unknown as Record<string, any>
    value.extra = true
    value.sections[0].page.margins.inside_twips = 200
    value.body.blocks[0].paragraph.runs[0].properties.css_class = 'heading'
    const decoded = decodeNativeDocxDocument(value)
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues.map(({ code, path }) => ({ code, path }))).toEqual(expect.arrayContaining([
      { code: 'UNKNOWN_FIELD', path: '/extra' },
      { code: 'UNKNOWN_FIELD', path: '/sections/0/page/margins/inside_twips' },
      { code: 'UNKNOWN_FIELD', path: '/body/blocks/0/paragraph/runs/0/properties/css_class' },
    ]))
  })

  it('rejects duplicate identities, dangling references, invalid unions, and unsafe edit policies', () => {
    const value = structuredClone(fixture) as unknown as Record<string, any>
    value.headers[0].id = 'story:body'
    value.sections[0].starts_at_block_id = 'paragraph:missing'
    value.body.blocks[0].paragraph.runs[0].control = 'tab'
    value.body.blocks[1].table.edit_policy.refusal = undefined
    const decoded = decodeNativeDocxDocument(value)
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues.map((entry) => entry.code)).toEqual(expect.arrayContaining([
      'DUPLICATE_ID',
      'BROKEN_REFERENCE',
      'INVALID_UNION',
      'REQUIRED',
    ]))
  })

  it('requires exact relationship-bound decimal note roles and unique native ids', () => {
    const duplicate = structuredClone(fixture) as unknown as Record<string, any>
    const copied = structuredClone(duplicate.notes[0])
    copied.id = 'note:footnote:duplicate'
    copied.blocks = []
    duplicate.notes.push(copied)
    const duplicateResult = decodeNativeDocxDocument(duplicate)
    expect(duplicateResult).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'DUPLICATE_ID', path: '/notes/1/native_story_id' })]) }))

    const relationship = structuredClone(fixture) as any
    const endnote = structuredClone(relationship.notes[0])
    endnote.id = 'note:endnote:1'; endnote.kind = 'endnote'; endnote.part_name = 'word/endnotes.xml'; endnote.anchor.part_name = endnote.part_name; endnote.blocks = []
    relationship.notes.push(endnote)
    expect(decodeNativeDocxDocument(relationship)).toEqual(expect.objectContaining({ ok: false, issues: expect.arrayContaining([expect.objectContaining({ code: 'DUPLICATE_ID', path: '/notes/1/relationship_id' })]) }))

    for (const mutate of [
      (value: any) => { value.notes[0].native_story_id = 'abc' },
      (value: any) => { value.notes[0].native_story_id = '0' },
      (value: any) => { delete value.notes[0].relationship_id },
      (value: any) => { delete value.notes[0].note_role },
      (value: any) => { value.body.note_role = 'content' },
    ]) {
      const value = structuredClone(fixture) as any
      mutate(value)
      expect(decodeNativeDocxDocument(value).ok).toBe(false)
    }
  })

  it('rejects malformed anchors and fingerprints', () => {
    const value = structuredClone(fixture) as unknown as Record<string, any>
    value.source.package_sha256 = 'not-a-hash'
    value.body.anchor.part_name = '../document.xml'
    value.body.anchor.end_byte = 10
    const decoded = decodeNativeDocxDocument(value)
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues.map((entry) => entry.path)).toEqual(expect.arrayContaining([
      '/source/package_sha256',
      '/body/anchor/part_name',
      '/body/anchor/end_byte',
    ]))
  })

  it('throws a structured error rather than serializing an invalid typed value', () => {
    const value = structuredClone(fixture)
    value.version = 2 as 1
    expect(() => encodeNativeDocxDocument(value)).toThrow(NativeDocxValidationError)
  })

  it('keeps TypeScript binding fields aligned with the published canonical schema', () => {
    const schema = JSON.parse(readFileSync(new URL('../schema/native-docx-v1.schema.json', import.meta.url), 'utf8')) as {
      $defs: Record<string, { properties?: Record<string, unknown> }>
    }
    for (const [name, fields] of Object.entries(DOCX_NATIVE_V1_BINDING_FIELDS)) {
      expect(Object.keys(schema.$defs[name]?.properties ?? {}).sort(), name).toEqual([...fields].sort())
    }
  })

  it.each(invalidFixture.cases)('rejects shared invalid vector: $name', ({ pointer, value, code, path }) => {
    const candidate = structuredClone(fixture) as unknown as Record<string, any>
    setPointer(candidate, pointer, value)
    const decoded = decodeNativeDocxDocument(candidate)
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues).toEqual(expect.arrayContaining([expect.objectContaining({ code, path })]))
  })

  it('bounds raw payloads, collection traversal, text, and issue output', () => {
    const hugePayload = ' '.repeat(DOCX_NATIVE_LIMITS.maxJsonBytes + 1)
    expect(decodeNativeDocxJson(hugePayload)).toEqual(expect.objectContaining({ ok: false }))

    const tooManyOperations = structuredClone(fixture) as unknown as Record<string, any>
    tooManyOperations.body.blocks[0].paragraph.edit_policy.allowed_operations = [
      'text.replace', 'properties.patch', 'block.insert_after', 'block.delete', 'drawing.replace', 'text.replace',
    ]
    const operationsResult = decodeNativeDocxDocument(tooManyOperations)
    expect(operationsResult.ok).toBe(false)
    if (!operationsResult.ok) expect(operationsResult.issues).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'LIMIT_EXCEEDED', path: '/body/blocks/0/paragraph/edit_policy/allowed_operations' })]))

    const tooMuchText = structuredClone(fixture) as unknown as Record<string, any>
    tooMuchText.body.blocks[0].paragraph.runs[0].text = 'x'.repeat(DOCX_NATIVE_LIMITS.maxTextLength + 1)
    const textResult = decodeNativeDocxDocument(tooMuchText)
    expect(textResult.ok).toBe(false)
    if (!textResult.ok) expect(textResult.issues).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'LIMIT_EXCEEDED', path: '/body/blocks/0/paragraph/runs/0/text' })]))

    const manyIssues = structuredClone(fixture) as unknown as Record<string, any>
    for (let index = 0; index < 200; index += 1) manyIssues[`unknown_${index}`] = true
    const issueResult = decodeNativeDocxDocument(manyIssues)
    expect(issueResult.ok).toBe(false)
    if (!issueResult.ok) expect(issueResult.issues).toHaveLength(DOCX_NATIVE_LIMITS.maxIssues)
  })

  it('rejects raw negative zero before canonicalization can erase it', () => {
    const raw = readFileSync(new URL('../../../testdata/docx-native/document-v1.json', import.meta.url), 'utf8')
      .replace('"start_byte": 100,', '"start_byte": -0,')
    const decoded = decodeNativeDocxJson(raw)
    expect(decoded.ok).toBe(false)
    if (!decoded.ok) expect(decoded.issues).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'INVALID_VALUE', path: '/body/anchor/start_byte' })]))
  })

  it('publishes structural bounds and exclusive unions in the canonical schema', () => {
    const schema = JSON.parse(readFileSync(new URL('../schema/native-docx-v1.schema.json', import.meta.url), 'utf8')) as any
    expect(schema.$defs.DocumentV1.properties.comment_stories.maxItems).toBe(DOCX_NATIVE_LIMITS.maxCollectionItems)
    expect(schema.$defs.RunV1.properties.text.maxLength).toBe(DOCX_NATIVE_LIMITS.maxTextLength)
    expect(schema.$defs.RunV1.oneOf.every((branch: any) => branch.not)).toBe(true)
    expect(schema.$defs.BlockV1.oneOf.every((branch: any) => branch.not)).toBe(true)
    expect(schema.$defs.StoryV1.allOf[0].then.required).toEqual(expect.arrayContaining(['native_story_id', 'relationship_id', 'note_role']))
    expect(schema.$defs.StoryV1.allOf[0].then.oneOf.map((branch: any) => branch.properties.note_role.const)).toEqual(['content', 'separator', 'continuation-separator'])
    expect(schema.$defs.PartName.pattern).toEqual(expect.stringContaining('2E|2F'))
    expect(schema.$defs.PartName.pattern).toEqual(expect.stringContaining('5C'))
    expect(schema.$defs.PartName.pattern).toEqual(expect.stringContaining('\\.(?:/|$)'))
  })
})
