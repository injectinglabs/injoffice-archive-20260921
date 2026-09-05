import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import type { NativeDocxDocumentV1, NativeDocxParagraphV1 } from '../../../packages/docs/src/nativeContract'
import {
  DOCX_PREVIEW_BLOCK_LIMIT,
  nativeDocxParagraphText,
  nativeDocxPreviewStats,
  nativeDocxStoryText,
  visibleNativeDocxBlocks,
} from './docsNativePreview'

const anchor = { part_name: 'word/document.xml', path: '/w:document/w:body/w:p[1]', start_byte: 1, end_byte: 2, xml_sha256: `sha256:${'a'.repeat(64)}` }
const editPolicy = { mode: 'read-write' as const, allowed_operations: ['text.replace' as const] }

const paragraph: NativeDocxParagraphV1 = {
  id: 'paragraph:1', anchor, edit_policy: editPolicy, properties: {},
  runs: [
    { kind: 'text', id: 'run:1', anchor, text: 'Hello' },
    { kind: 'control', id: 'run:2', anchor, control: 'tab' },
    { kind: 'text', id: 'run:3', anchor, text: 'hidden', properties: { hidden: true } },
    { kind: 'text', id: 'run:4', anchor, text: 'world' },
    { kind: 'drawing', id: 'run:5', anchor, drawing: { id: 'drawing:1', anchor, placement: 'inline', width_emu: 1, height_emu: 1, edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'READ_ONLY', message: 'bounded', preservation: 'refuse-mutation' } } } },
  ],
}

function fixture(blocks = [{ kind: 'paragraph' as const, id: 'block:1', paragraph }]): NativeDocxDocumentV1 {
  return {
    protocol: 'injoffice.docx.native', version: 1, document_id: 'doc:1', revision: 'revision:1',
    source: { package_sha256: `sha256:${'b'.repeat(64)}`, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor, blocks },
    sections: [], headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [], passthrough_parts: [], unsupported: [],
  }
}

describe('native DOCX playground proof', () => {
  it('ships a decodable, substantive DOCX package for the native extraction path', () => {
    const encoded = readFileSync(new URL('../public/native-docx/northstar-launch-brief.docx.b64', import.meta.url), 'utf8').trim()
    const bytes = Buffer.from(encoded, 'base64')
    expect(bytes.subarray(0, 4).toString('hex')).toBe('504b0304')
    expect(bytes.byteLength).toBeGreaterThan(1_500)
    expect(encoded.length).toBeGreaterThan(2_000)
  })

  it('projects modeled visible text without leaking hidden runs', () => {
    expect(nativeDocxParagraphText(paragraph)).toBe('Hello\tworld')
    expect(nativeDocxStoryText(fixture().body)).toBe('Hello\tworld')
  })

  it('reports native blocks and inline structures instead of invented pages', () => {
    expect(nativeDocxPreviewStats(fixture())).toEqual({ blocks: 1, paragraphs: 1, tables: 0, textRuns: 3, drawings: 1, references: 0, readOnlyBlocks: 0 })
  })

  it('bounds browser rendering while retaining an honest omitted count', () => {
    const blocks = Array.from({ length: DOCX_PREVIEW_BLOCK_LIMIT + 7 }, (_, index) => ({ kind: 'paragraph' as const, id: `block:${index}`, paragraph: { ...paragraph, id: `paragraph:${index}` } }))
    const preview = visibleNativeDocxBlocks(fixture(blocks))
    expect(preview.blocks).toHaveLength(DOCX_PREVIEW_BLOCK_LIMIT)
    expect(preview.omitted).toBe(7)
  })
})
