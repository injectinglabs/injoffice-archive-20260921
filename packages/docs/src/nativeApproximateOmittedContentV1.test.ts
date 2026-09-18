import { describe, expect, it } from 'vitest'
import { collectNativeDocxApproximateOmissionsV1, nativeDocxOmittedContentCategoryV1, nativeDocxOmittedContentSummaryV1, nativeDocxApproximateRefusalOmissionsV1, validNativeDocxApproximateOmissionsV1, DOCX_APPROXIMATE_OMITTED_CONTENT_CODES, DOCX_APPROXIMATE_OMITTED_CONTENT_LIMIT, type NativeDocxApproximateOmissionSourceV1 } from './nativeApproximateOmittedContentV1.js'
import { DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED } from './nativeApproximationV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const anchor = (path: string) => ({ part_name: 'word/document.xml', path, start_byte: 1, end_byte: 2, xml_sha256: HASH })

function source(overrides: { runs?: any[]; unsupported?: any[]; resolution?: any[]; shaping?: any[]; shapedParagraphs?: string[] } = {}): NativeDocxApproximateOmissionSourceV1 {
  const runs = overrides.runs ?? [{ kind: 'text', id: 'run:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]'), text: 'AA' }]
  const paragraph = { id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]'), edit_policy: {}, properties: {}, runs }
  return {
    document: { document_id: 'document:test', body: { id: 'story:body', blocks: [{ kind: 'paragraph', id: paragraph.id, paragraph }] }, headers: [], footers: [], notes: [], comment_stories: [], unsupported: overrides.unsupported ?? [] } as any,
    resolved_layout: { document_id: 'document:test', source_parts: { main_part: 'word/document.xml' }, diagnostics: overrides.resolution ?? [] } as any,
    shaped_lines: { paragraphs: (overrides.shapedParagraphs ?? ['paragraph:1']).map((id) => ({ paragraph_id: id })), diagnostics: overrides.shaping ?? [] } as any,
  }
}
const page = (id: string, commands: number) => ({ id, commands: Array.from({ length: commands }, (_, index) => ({ kind: 'fill_glyph_path', id: `${id}:${index}` })) }) as any

describe('approximate omitted-content disclosure', () => {
  it('reports complete when every source item was painted', () => {
    const result = collectNativeDocxApproximateOmissionsV1(source(), { status: 'painted', pages: [page('page:1', 2)] })
    expect(result).toEqual({ content_status: 'complete', omitted_content: [], omitted_content_total: 0, unpainted_pages: [] })
    expect(validNativeDocxApproximateOmissionsV1(result, { status: 'painted', pages: [page('page:1', 2)] })).toBe(true)
    expect(nativeDocxOmittedContentSummaryV1(result)).toBeNull()
  })
  it('discloses a rotated table cell as table content the preview did not render', () => {
    // w:tcPr/w:textDirection rotates the cell's text 90 or 270 degrees. The
    // cell's glyphs are still painted, horizontally, so the code has to be in
    // both sets: the paint tier omits it instead of refusing, and this
    // disclosure reports it so the page can never claim to be complete.
    expect(DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED.has('CELL_TEXT_DIRECTION_UNSUPPORTED')).toBe(true)
    expect(DOCX_APPROXIMATE_OMITTED_CONTENT_CODES.has('CELL_TEXT_DIRECTION_UNSUPPORTED')).toBe(true)
    const path = '/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[2]/w:tcPr[1]/w:textDirection[1]'
    const unsupported = [{ id: 'unsupported:1', code: 'CELL_TEXT_DIRECTION_UNSUPPORTED', capability: 'table-properties', scope_id: 'table:1', anchor: anchor(path), preservation: 'refuse-mutation', message: 'Rotated or vertically stacked cell text direction btLr is recorded and not applied; the cell\u2019s text is laid out and painted horizontally' }]
    const result = collectNativeDocxApproximateOmissionsV1(source({ unsupported }), { status: 'painted', pages: [page('page:1', 2)] })
    expect(result.content_status).toBe('partial')
    expect(result.omitted_content).toEqual([expect.objectContaining({ code: 'CELL_TEXT_DIRECTION_UNSUPPORTED', origin: 'source', category: 'table', scope_id: 'table:1', path, count: 1 })])
    expect(nativeDocxOmittedContentSummaryV1(result)).toBe('1 item not rendered: tables (1)')
    expect(validNativeDocxApproximateOmissionsV1(result, { status: 'painted', pages: [page('page:1', 2)] })).toBe(true)
  })

  it('discloses a drawing-only document as partial with one drawing entry and the blank page', () => {
    const drawing = { kind: 'drawing', id: 'run:d', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]'), drawing: { id: 'drawing:1' } }
    const unsupported = [{ id: 'unsupported:1', code: 'UNMODELED_DRAWING', capability: 'drawings', scope_id: 'run:d', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]'), preservation: 'refuse-mutation', message: 'Drawing/object markup and related media are preserved verbatim' }]
    const result = collectNativeDocxApproximateOmissionsV1(source({ runs: [drawing], unsupported, shapedParagraphs: [] }), { status: 'painted', pages: [page('page:1', 0)] })
    expect(result.content_status).toBe('partial')
    expect(result.unpainted_pages).toEqual(['page:1'])
    expect(result.omitted_content.filter((entry) => entry.code === 'UNMODELED_DRAWING')).toEqual([{ code: 'UNMODELED_DRAWING', origin: 'source', category: 'drawing', scope_id: 'run:d', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', message: 'Drawing/object markup and related media are preserved verbatim', count: 1 }])
    expect(result.omitted_content.map((entry) => entry.code)).toEqual(['UNMODELED_DRAWING'])
    expect(result.omitted_content_total).toBe(1)
    expect(nativeDocxOmittedContentSummaryV1(result)).toBe('1 item not rendered: drawings (1); 1 page painted nothing')
    const unexplained = collectNativeDocxApproximateOmissionsV1(source({ runs: [drawing], shapedParagraphs: [] }), { status: 'painted', pages: [page('page:1', 0)] })
    expect(unexplained.omitted_content).toEqual([expect.objectContaining({ code: 'PARAGRAPH_NOT_SHAPED', origin: 'pagination', category: 'drawing', scope_id: 'paragraph:1', count: 1 })])
    expect(nativeDocxOmittedContentSummaryV1(unexplained)).toBe('1 item not rendered: drawings (1); 1 page painted nothing')
    expect(validNativeDocxApproximateOmissionsV1(result, { status: 'painted', pages: [page('page:1', 0)] })).toBe(true)
    expect(validNativeDocxApproximateOmissionsV1({ ...result, content_status: 'complete' }, { status: 'painted', pages: [page('page:1', 0)] })).toBe(false)
    expect(validNativeDocxApproximateOmissionsV1({ ...result, unpainted_pages: ['page:2'] }, { status: 'painted', pages: [page('page:1', 0)] })).toBe(false)
    expect(validNativeDocxApproximateOmissionsV1({ ...result, omitted_content_total: 0 }, { status: 'painted', pages: [page('page:1', 0)] })).toBe(false)
  })
  it('keeps formatting-only approximations out of the omitted list and merges duplicates', () => {
    const resolution = [
      { code: 'THEME_COLOR_PRESERVED', severity: 'unsupported', scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]/w:color[1]', preservation: 'preserve-verbatim', message: 'kept' },
      { code: 'UNMODELED_RUN_CONTENT', severity: 'unsupported', scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/m:oMath[1]', preservation: 'preserve-verbatim', message: 'math' },
      { code: 'UNMODELED_RUN_CONTENT', severity: 'unsupported', scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/m:oMath[1]', preservation: 'preserve-verbatim', message: 'math' },
    ]
    const shaping = [
      { code: 'drawing-layout-unsupported', severity: 'unsupported', scope_id: 'paragraph:1', source_id: 'run:1', message: 'Drawing payload is missing' },
      { code: 'page-control-deferred', severity: 'deferred', scope_id: 'paragraph:1', message: 'deferred' },
    ]
    const result = collectNativeDocxApproximateOmissionsV1(source({ resolution, shaping }), { status: 'painted', pages: [page('page:1', 1)] })
    expect(result.omitted_content).toEqual([
      { code: 'drawing-layout-unsupported', origin: 'shaping', category: 'drawing', scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]', message: 'Drawing payload is missing', count: 1 },
      { code: 'UNMODELED_RUN_CONTENT', origin: 'resolution', category: 'equation', scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/m:oMath[1]', message: 'math', count: 2 },
    ])
    expect(result).toMatchObject({ content_status: 'partial', omitted_content_total: 3, unpainted_pages: [] })
    expect(nativeDocxOmittedContentSummaryV1(result)).toBe('3 items not rendered: equations (2), drawings (1)')
  })
  it('bounds the list, keeps the total, and still flags a blank page from an empty source', () => {
    const unsupported = Array.from({ length: 80 }, (_, index) => ({ id: `u:${index}`, code: 'UNMODELED_BODY_BLOCK', capability: 'body', scope_id: `block:${index}`, anchor: anchor(`/w:document[1]/w:body[1]/w:sdt[${index + 1}]`), preservation: 'preserve-verbatim', message: 'block' }))
    const bounded = collectNativeDocxApproximateOmissionsV1(source({ unsupported }), { status: 'painted', pages: [page('page:1', 1)] })
    expect(bounded.omitted_content).toHaveLength(DOCX_APPROXIMATE_OMITTED_CONTENT_LIMIT)
    expect(bounded.omitted_content_total).toBe(80)
    expect(nativeDocxOmittedContentSummaryV1(bounded)).toBe('80 items not rendered: content controls (64), 16 more not listed')
    expect(validNativeDocxApproximateOmissionsV1(bounded, { status: 'painted', pages: [page('page:1', 1)] })).toBe(true)
    const empty = collectNativeDocxApproximateOmissionsV1(source({ runs: [] }), { status: 'painted', pages: [page('page:1', 0), page('page:2', 1)] })
    expect(empty).toEqual({ content_status: 'partial', omitted_content: [], omitted_content_total: 0, unpainted_pages: ['page:1'] })
    expect(validNativeDocxApproximateOmissionsV1({ ...empty, unpainted_pages: [] }, { status: 'painted', pages: [page('page:1', 0), page('page:2', 1)] })).toBe(false)
    expect(validNativeDocxApproximateOmissionsV1({ ...empty, unpainted_pages: [], content_status: 'complete' }, { status: 'painted', pages: [page('page:1', 0), page('page:2', 1)] })).toBe(false)
  })
  it('does not count non-visual markers as omitted content', () => {
    const unsupported = [
      { id: 'u:1', code: 'UNMODELED_PARAGRAPH_CONTENT', capability: 'paragraphs', scope_id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:bookmarkStart[1]'), preservation: 'preserve-verbatim', message: 'bookmark' },
      { id: 'u:2', code: 'UNMODELED_PARAGRAPH_CONTENT', capability: 'paragraphs', scope_id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:bookmarkEnd[1]'), preservation: 'preserve-verbatim', message: 'bookmark' },
      { id: 'u:3', code: 'UNMODELED_PARAGRAPH_CONTENT', capability: 'paragraphs', scope_id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:proofErr[1]'), preservation: 'preserve-verbatim', message: 'proofing' },
    ]
    const resolution = [{ code: 'UNMODELED_PARAGRAPH_CONTENT', severity: 'unsupported', scope_id: 'paragraph:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:permStart[1]', preservation: 'preserve-verbatim', message: 'permission' }]
    const result = collectNativeDocxApproximateOmissionsV1(source({ unsupported, resolution }), { status: 'painted', pages: [page('page:1', 2)] })
    expect(result).toEqual({ content_status: 'complete', omitted_content: [], omitted_content_total: 0, unpainted_pages: [] })
  })
  it('treats refusals as a fixed partial record', () => {
    const refusal = nativeDocxApproximateRefusalOmissionsV1()
    expect(validNativeDocxApproximateOmissionsV1(refusal, { status: 'refused', pages: [] })).toBe(true)
    expect(validNativeDocxApproximateOmissionsV1({ ...refusal, content_status: 'complete' }, { status: 'refused', pages: [] })).toBe(false)
    expect(collectNativeDocxApproximateOmissionsV1(source(), { status: 'refused', pages: [] })).toEqual(refusal)
  })
  it('categorizes by source path before code family', () => {
    expect(nativeDocxOmittedContentCategoryV1('UNMODELED_PARAGRAPH_CONTENT', '/w:document[1]/w:body[1]/w:p[1]/m:oMathPara[1]')).toBe('equation')
    expect(nativeDocxOmittedContentCategoryV1('UNMODELED_RUN_CONTENT', '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:pict[1]')).toBe('drawing')
    expect(nativeDocxOmittedContentCategoryV1('UNMODELED_BODY_BLOCK', '/w:document[1]/w:body[1]/w:sdt[1]')).toBe('content-control')
    expect(nativeDocxOmittedContentCategoryV1('UNMODELED_BODY_BLOCK', undefined)).toBe('block')
    expect(nativeDocxOmittedContentCategoryV1('FIELD_SEMANTICS', '/w:document[1]/w:body[1]/w:p[1]/w:r[1]')).toBe('field')
    expect(nativeDocxOmittedContentCategoryV1('WRAPPED_RUN_MARKUP', '/w:document[1]/w:body[1]/w:p[1]/w:ins[1]')).toBe('revision')
    expect(nativeDocxOmittedContentCategoryV1('reference-layout-unsupported', undefined)).toBe('reference')
    expect(nativeDocxOmittedContentCategoryV1('SOMETHING_ELSE', undefined)).toBe('other')
    expect(nativeDocxOmittedContentCategoryV1('UNMODELED_RUN_CONTENT', '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/ns4d2aa588:AlternateContent[1]')).toBe('drawing')
    expect(nativeDocxOmittedContentCategoryV1('PICTURE_GRAPHIC_REQUIRED', '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]/ns4a8a39ac:inline[1]')).toBe('drawing')
    expect(nativeDocxOmittedContentCategoryV1('PICTURE_TRANSFORM_PRESERVED', undefined)).toBe('drawing')
    expect(nativeDocxOmittedContentCategoryV1('table-layout-unsupported', undefined)).toBe('table')
    expect(nativeDocxOmittedContentCategoryV1('unsupported-numbering-text', undefined)).toBe('text')
    expect(nativeDocxOmittedContentCategoryV1('UNMODELED_PARAGRAPH_CONTENT', '/w:document[1]/w:body[1]/w:p[1]/ns1234abcd:oMathPara[1]')).toBe('equation')
    expect(nativeDocxOmittedContentCategoryV1('COLUMN_SEPARATOR_UNSUPPORTED', '/w:document[1]/w:body[1]/w:p[2]/w:pPr[1]/w:sectPr[1]/w:cols[1]')).toBe('other')
    for (const code of ['STYLISTIC_SET_UNAPPLIED', 'LIGATURE_MODE_UNAPPLIED', 'NUMBER_FORM_UNAPPLIED', 'NUMBER_SPACING_UNAPPLIED', 'TEXT_EFFECT_3D_UNAPPLIED']) {
      expect(nativeDocxOmittedContentCategoryV1(code, '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]/nsdbbea4e0:numForm[1]'), code).toBe('text')
    }
  })
  it('discloses every run-typography feature the approximate tier paints without', () => {
    // The approximate tier paints these runs with the feature unapplied, so the
    // requested glyph forms and advances are absent from the page. That is a
    // dropped visible mark, and this is the discloser that names it; the
    // invariant below pins that the paint set may not admit one of these codes
    // without this list reporting it.
    const codes = ['STYLISTIC_SET_UNAPPLIED', 'LIGATURE_MODE_UNAPPLIED', 'NUMBER_FORM_UNAPPLIED', 'NUMBER_SPACING_UNAPPLIED', 'TEXT_EFFECT_3D_UNAPPLIED'] as const
    for (const code of codes) {
      const unsupported = [{ id: 'u:1', code, capability: 'run-properties', scope_id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]/nsdbbea4e0:numForm[1]'), preservation: 'preserve-verbatim', message: `${code} is preserved and NOT applied` }]
      const result = collectNativeDocxApproximateOmissionsV1(source({ unsupported }), { status: 'painted', pages: [page('page:1', 2)] })
      expect(result, code).toMatchObject({ content_status: 'partial', omitted_content_total: 1 })
      expect(result.omitted_content, code).toEqual([expect.objectContaining({ code, origin: 'source', category: 'text', scope_id: 'paragraph:1', count: 1 })])
    }
    // The property this tier does apply is not an omission: a ligature mode the
    // declared shaper defaults already perform states no missing mark.
    const applied = [{ id: 'u:1', code: 'LIGATURE_MODE_MATCHES_SHAPER', capability: 'run-properties', scope_id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]/nsdbbea4e0:ligatures[1]'), preservation: 'preserve-verbatim', message: 'already applied' }]
    expect(collectNativeDocxApproximateOmissionsV1(source({ unsupported: applied }), { status: 'painted', pages: [page('page:1', 2)] })).toMatchObject({ content_status: 'complete', omitted_content: [], omitted_content_total: 0 })
  })
  it('never lets the paint set admit a run-typography code the discloser does not report', () => {
    // The invariant PR #334 documented, applied to this slice: a code the
    // approximate tier lays out around must either leave the painted content
    // intact or be named here, or the drop is silent.
    for (const code of ['STYLISTIC_SET_UNAPPLIED', 'LIGATURE_MODE_UNAPPLIED', 'NUMBER_FORM_UNAPPLIED', 'NUMBER_SPACING_UNAPPLIED', 'TEXT_EFFECT_3D_UNAPPLIED']) {
      expect(DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED.has(code), code).toBe(true)
      expect(DOCX_APPROXIMATE_OMITTED_CONTENT_CODES.has(code), code).toBe(true)
    }
    // Every other foreign run property stays outside both sets and keeps refusing.
    for (const code of ['FOREIGN_RUN_PROPERTY', 'UNMODELED_RUN_PROPERTY']) {
      expect(DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED.has(code), code).toBe(false)
      expect(DOCX_APPROXIMATE_OMITTED_CONTENT_CODES.has(code), code).toBe(false)
    }
  })
  it('records the column separator rule as dropped ink, not an approximated property', () => {
    // The rule w:cols w:sep asks for is ink Word draws in the inter-column gap
    // and this tier does not, so it belongs in omitted_content; the section
    // property records that accompany it stay formatting-only.
    const unsupported = [
      { id: 'u:1', code: 'COLUMN_SEPARATOR_UNSUPPORTED', capability: 'sections', scope_id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[2]/w:pPr[1]/w:sectPr[1]/w:cols[1]'), preservation: 'preserve-verbatim', message: 'Column separators require paint geometry that v1 does not model' },
      { id: 'u:2', code: 'UNMODELED_SECTION_PROPERTY', capability: 'sections', scope_id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]/w:docGrid[1]'), preservation: 'preserve-verbatim', message: 'grid' },
    ]
    const result = collectNativeDocxApproximateOmissionsV1(source({ unsupported }), { status: 'painted', pages: [page('page:1', 2)] })
    expect(result).toMatchObject({ content_status: 'partial', omitted_content_total: 1 })
    expect(result.omitted_content).toEqual([expect.objectContaining({ code: 'COLUMN_SEPARATOR_UNSUPPORTED', origin: 'source', category: 'other', scope_id: 'section:1', count: 1 })])
  })
})
