import { describe, expect, it } from 'vitest'
import {
  DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL,
  DOCX_HEADER_FOOTER_LAYOUT_VERSION,
  layoutNativeDocxHeadersFootersV1,
  nativeDocxHeaderFooterLayoutSha256V1,
  type NativeDocxHeaderFooterLayoutInputV1,
  type NativeDocxHeaderFooterLayoutSuccessV1,
} from './nativeHeaderFooterLayoutV1.js'

const HASH = `sha256:${'a'.repeat(64)}`

function anchor(part: string, path: string) {
  return { part_name: part, path, start_byte: 1, end_byte: 2, xml_sha256: HASH }
}

function paragraph(id: string, part: string) {
  return {
    id, anchor: anchor(part, `/w:hdr[1]/w:p[${id}]`),
    edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'READ_ONLY', message: 'fixture', preservation: 'refuse-mutation' } },
    properties: {}, runs: [],
  }
}

function story(id: string, kind: 'header' | 'footer') {
  const part = `word/${id}.xml`
  const value = paragraph(`paragraph:${id}`, part)
  return { id: `story:${id}`, kind, part_name: part, anchor: anchor(part, kind === 'header' ? '/w:hdr[1]' : '/w:ftr[1]'), blocks: [{ kind: 'paragraph', id: value.id, paragraph: value }] }
}

function shaped(stories: ReturnType<typeof story>[]) {
  return stories.map((value) => ({
    paragraph_id: value.blocks[0]!.paragraph.id, story_id: value.id, story_kind: value.kind,
    direction: 'ltr', alignment: 'start', spacing_before_millipoints: 0, spacing_after_millipoints: 0,
    indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0,
    block_advance_millipoints: 10_000,
    lines: [{
      id: `line:${value.id}`, ordinal: 0, available_width_millipoints: 468_000, inline_offset_millipoints: 0,
      advance_inline_millipoints: 20_000, ascent_millipoints: 8_000, descent_millipoints: -2_000,
      line_gap_millipoints: 0, line_height_millipoints: 10_000, fragments: [],
    }],
  }))
}

function fixture(): NativeDocxHeaderFooterLayoutInputV1 {
  const stories = [story('header-default', 'header'), story('header-first', 'header'), story('header-even', 'header'), story('header-second-default', 'header'), story('footer-default', 'footer')]
  const bodyParagraph = paragraph('paragraph:body', 'word/document.xml')
  const page = {
    width_twips: 12_240, height_twips: 15_840, orientation: 'portrait' as const,
    margins: { top_twips: 1_440, right_twips: 1_440, bottom_twips: 1_440, left_twips: 1_440, header_twips: 720, footer_twips: 720, gutter_twips: 0 },
    columns: 1, column_spacing_twips: 720,
  }
  const ref = (kind: 'default' | 'first' | 'even', id: string) => ({ kind, story_id: `story:${id}`, relationship_id: `r:${id}` })
  const sections = [{
    id: 'section:1', anchor: anchor('word/document.xml', '/w:document[1]/w:body[1]/w:p[1]/w:pPr[1]/w:sectPr[1]'), starts_at_block_id: bodyParagraph.id,
    break_type: 'next-page' as const, title_page: true, page,
    header_refs: [ref('default', 'header-default'), ref('first', 'header-first'), ref('even', 'header-even')],
    footer_refs: [ref('default', 'footer-default'), ref('first', 'footer-default'), ref('even', 'footer-default')],
  }, {
    id: 'section:2', anchor: anchor('word/document.xml', '/w:document[1]/w:body[1]/w:sectPr[1]'), starts_at_block_id: bodyParagraph.id,
    break_type: 'next-page' as const, title_page: false, page,
    header_refs: [ref('default', 'header-second-default')], footer_refs: [],
  }]
  const makePage = (id: string, ordinal: number, sectionID: string, sectionOrdinal: number) => ({
    id, ordinal, section_id: sectionID, section_ids: [sectionID], section_page_ordinal: sectionOrdinal, kind: 'content' as const,
    width_millipoints: 612_000, height_millipoints: 792_000,
    body_box: { x_millipoints: 72_000, y_millipoints: 72_000, width_millipoints: 468_000, height_millipoints: 648_000 },
    columns: [{ id: `${sectionID}:column:0`, section_id: sectionID, ordinal: 0, x_millipoints: 72_000, y_millipoints: 72_000, width_millipoints: 468_000, height_millipoints: 648_000 }],
    header_refs: [], footer_refs: [], paragraph_slices: [], lines: [],
  })
  return {
    document: {
      protocol: 'injoffice.docx.native', version: 1, document_id: 'document:test', revision: 'revision:1', source: { package_sha256: HASH, main_part: 'word/document.xml' },
      body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('word/document.xml', '/w:document[1]/w:body[1]'), blocks: [{ kind: 'paragraph', id: bodyParagraph.id, paragraph: bodyParagraph }] },
      sections, headers: stories.filter((value) => value.kind === 'header'), footers: stories.filter((value) => value.kind === 'footer'), notes: [], comment_stories: [], comments: [], capabilities: [], passthrough_parts: [], unsupported: [],
    } as any,
    resolved_layout: {
      protocol: 'injoffice.docx.resolved-layout', version: 1, document_id: 'document:test', revision: 'revision:1', source_parts: { main_part: 'word/document.xml' },
      paragraphs: stories.map((value) => ({ paragraph_id: value.blocks[0]!.paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: {} })),
      runs: [], tables: [], fonts: [], diagnostics: [],
    } as any,
    shaped_lines: {
      protocol: 'injoffice.docx.shaped-lines', version: 1, document_id: 'document:test', revision: 'revision:1', available_width_millipoints: 468_000, tab_interval_millipoints: 36_000,
      font_manifest: { manifest_id: 'manifest:test', revision: 'revision:1' }, providers: { resolver_id: 'resolver', resolver_revision: '1', shaper_id: 'shaper', shaper_revision: '1' }, paragraphs: shaped(stories), diagnostics: [],
    } as any,
    pagination_settings: {
      protocol: 'injoffice.docx.pagination-settings', version: 1, document_id: 'document:test', revision: 'revision:1', package_sha256: HASH, main_part: 'word/document.xml',
      profile: 'word-modern-default', default_tab_stop_twips: 720, mirror_margins: false, gutter_at_top: false, even_and_odd_headers: true, compatibility_mode: 15, diagnostics: [],
    },
    paginated_layout: {
      protocol: 'injoffice.docx.paginated-layout', version: 1, status: 'paginated', provenance: {} as any, diagnostics: [],
      sections: [], pages: [makePage('page:section:1:0', 0, 'section:1', 0), makePage('page:section:1:1', 1, 'section:1', 1), makePage('page:section:2:0', 2, 'section:2', 0), makePage('page:section:2:1', 3, 'section:2', 1)],
    },
  }
}

describe('native DOCX header/footer layout v1', () => {
  it('resolves independent inheritance and exact title/even/default physical-page selection', () => {
    const value = layoutNativeDocxHeadersFootersV1(fixture())
    expect(value.status).toBe('placed')
    if (value.status !== 'placed') throw new Error(JSON.stringify(value.diagnostics))
    expect(value.pages.map((page) => page.header_ref?.story_id)).toEqual(['story:header-first', 'story:header-even', 'story:header-second-default', 'story:header-even'])
    expect(value.pages.map((page) => page.footer_ref?.story_id)).toEqual(['story:footer-default', 'story:footer-default', 'story:footer-default', 'story:footer-default'])
    expect(value.pages[0]!.lines.find((line) => line.region === 'header')).toEqual(expect.objectContaining({ y_millipoints: 36_000, x_millipoints: 72_000 }))
    expect(value.pages[0]!.lines.find((line) => line.region === 'footer')).toEqual(expect.objectContaining({ y_millipoints: 746_000, x_millipoints: 72_000 }))
    expect(value.sha256).toBe(nativeDocxHeaderFooterLayoutSha256V1(value))
    expect(layoutNativeDocxHeadersFootersV1(structuredClone(fixture()))).toEqual(value)
  })

  it('hashes canonical object keys while preserving semantic array order', () => {
    const left: Omit<NativeDocxHeaderFooterLayoutSuccessV1, 'sha256'> = { protocol: DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL, version: DOCX_HEADER_FOOTER_LAYOUT_VERSION, status: 'placed', diagnostics: [], pages: [{ page_id: 'page:1', lines: [] }] }
    const right: Omit<NativeDocxHeaderFooterLayoutSuccessV1, 'sha256'> = { pages: [{ lines: [], page_id: 'page:1' }], diagnostics: [], status: 'placed', version: DOCX_HEADER_FOOTER_LAYOUT_VERSION, protocol: DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL }
    expect(nativeDocxHeaderFooterLayoutSha256V1(left)).toBe(nativeDocxHeaderFooterLayoutSha256V1(right))
    expect(nativeDocxHeaderFooterLayoutSha256V1({ ...left, pages: [{ page_id: 'page:2', lines: [] }, ...left.pages] })).not.toBe(nativeDocxHeaderFooterLayoutSha256V1(left))
  })

  it('refuses selected fields, tables, shapes, and one-unit reserved-band overflow without partial pages', () => {
    const cases: Array<[string, (value: NativeDocxHeaderFooterLayoutInputV1) => void]> = [
      ['selected-story-field', (value) => { value.document.unsupported.push({ id: 'unsupported:field', code: 'FIELD_SEMANTICS', capability: 'fields', scope_id: 'paragraph:header-first', preservation: 'refuse-mutation', message: 'PAGE field is not explicitly modeled' }) }],
      ['selected-story-table', (value) => { (value.document.headers.find((entry) => entry.id === 'story:header-first')!.blocks as any[]).push({ kind: 'table', id: 'table:1', table: { id: 'table:1', rows: [] } }) }],
      ['selected-story-shape', (value) => { (value.document.headers.find((entry) => entry.id === 'story:header-first')!.blocks[0]!.paragraph!.runs as any[]).push({ kind: 'drawing', id: 'run:drawing', drawing: { id: 'drawing:1' } }) }],
      ['header-band-overflow', (value) => { value.document.sections[0]!.page.margins.top_twips = 919; value.paginated_layout.status === 'paginated' && (value.paginated_layout.pages[0]!.body_box.y_millipoints = 45_950) }],
      ['section-geometry-invalid', (value) => { value.document.unsupported.push({ id: 'unsupported:section', code: 'UNMODELED_SECTION_PROPERTY', capability: 'sections', scope_id: 'section:1', preservation: 'refuse-mutation', message: 'Ambiguous section geometry' }) }],
    ]
    for (const [code, mutate] of cases) {
      const input = fixture(); mutate(input)
      const value = layoutNativeDocxHeadersFootersV1(input)
      expect(value).toEqual(expect.objectContaining({ status: 'refused', pages: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code })]) }))
    }
  })

  it('refuses inconsistent relationship-id reuse before returning any placement', () => {
    const input = fixture()
    input.document.sections[1]!.header_refs[0]!.relationship_id = 'r:header-even'
    const value = layoutNativeDocxHeadersFootersV1(input)
    expect(value).toEqual(expect.objectContaining({ status: 'refused', pages: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'relationship-ambiguous' })]) }))
  })
})
