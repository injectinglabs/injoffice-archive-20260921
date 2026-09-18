import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT,
  DOCX_ENCLOSED_MARKER_FONT_WARNING,
  projectNativeDocxEnclosedMarkerFontsV1,
  stripNativeDocxEnclosedMarkerFontsV1,
  validNativeDocxEnclosedMarkerFontsV1,
} from './nativeEnclosedMarkerFontV1.js'
import { DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED } from './nativeApproximationV1.js'
import { DOCX_APPROXIMATE_OMITTED_CONTENT_CODES, nativeDocxOmittedContentCategoryV1 } from './nativeApproximateOmittedContentV1.js'
import {
  decodeNativeDocxResolvedLayout,
  nativeDocxResolvedNumberingDefinitionSha256V1,
  nativeDocxResolvedNumberingModelSha256V1,
  type NativeDocxResolvedLayoutInputV1,
} from './nativeResolvedLayout.js'

const hash = `sha256:${'a'.repeat(64)}`
const fact = {
  scope_kind: 'numbering-marker' as const,
  scope_id: 'paragraph:1',
  part_name: 'word/document.xml',
  path: '/w:document[1]/w:body[1]/w:p[1]',
  source_family: 'Liberation Serif',
  package_sha256: hash,
}

describe('enclosed-number marker font evidence', () => {
  it('bounds unique source facts and rejects malformed or foreign fields', () => {
    expect(validNativeDocxEnclosedMarkerFontsV1([fact], hash)).toBe(true)
    for (const candidate of [
      [fact, fact],
      [{ ...fact, package_sha256: `sha256:${'b'.repeat(64)}` }],
      // A marker is the only scope this evidence has; no other scope kind exists.
      [{ ...fact, scope_kind: 'run' }],
      [{ ...fact, path: '/wrong' }],
      [{ ...fact, part_name: '../document.xml' }],
      // The source records the family it RESOLVED; it never carries the substitute.
      [{ ...fact, chosen_family: 'Cambria' }],
      [{ ...fact, source_family: ' Liberation Serif' }],
      [{ ...fact, source_family: '' }],
      Array.from({ length: 1001 }, (_, i) => ({ ...fact, scope_id: `paragraph:${i}` })),
    ]) expect(validNativeDocxEnclosedMarkerFontsV1(candidate, hash)).toBe(false)
  })

  it('states the substitute is a preview-host choice and names it', () => {
    expect(DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT).toBe('Cambria')
    expect(DOCX_ENCLOSED_MARKER_FONT_WARNING).toContain('Cambria')
    expect(DOCX_ENCLOSED_MARKER_FONT_WARNING).toContain('is not validated against Microsoft Word')
  })

  // The tested invariant: this set may only admit a content-dropping code that
  // the omitted-content discloser reports, so a substituted marker face can
  // never reach a page that calls itself complete.
  it('is a member of both approximation sets and is disclosed as text', () => {
    expect(DOCX_APPROXIMATE_OMITTED_SOURCE_UNSUPPORTED.has('ENCLOSED_NUMBER_MARKER_FONT_PRESERVED')).toBe(true)
    expect(DOCX_APPROXIMATE_OMITTED_CONTENT_CODES.has('ENCLOSED_NUMBER_MARKER_FONT_PRESERVED')).toBe(true)
    expect(nativeDocxOmittedContentCategoryV1('ENCLOSED_NUMBER_MARKER_FONT_PRESERVED', undefined)).toBe('text')
  })
})

describe('enclosed-number marker font projection', () => {
  const source = JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json', import.meta.url), 'utf8'))
  const relationshipsPart = 'word/_rels/document.xml.rels'
  const numberingPart = 'word/numbering.xml'
  const digest = (byte: string) => `sha256:${byte.repeat(64)}`

  function model(markerFamily = 'Liberation Serif') {
    const document = structuredClone(source)
    const paragraph = document.body.blocks[0].paragraph
    document.passthrough_parts = [
      ...document.passthrough_parts,
      { part_name: relationshipsPart, content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: digest('1'), policy: 'preserve-verbatim' },
      { part_name: numberingPart, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', byte_length: 1, sha256: digest('2'), policy: 'preserve-verbatim' },
    ]
    const numbering = {
      marker_id: 'marker:paragraph:intro', definition_sha256: digest('0'),
      num_id: '1', abstract_num_id: '1', level: 0, start: 1,
      format: 'decimalEnclosedCircle', text: '%1.', suffix: 'tab', alignment: 'left', never_restart: true,
      counter_value: 1, counter_values: [{ level: 0, value: 1, format: 'decimalEnclosedCircle' }], resolved_text: '\u2460.',
      label_start_twips: 360, label_end_twips: 720, text_start_twips: 720, numbering_tab_twips: 720,
      marker_properties: { font_family: markerFamily, font_size_half_points: 24 },
    }
    const resolved = {
      protocol: 'injoffice.docx.resolved-layout', version: 1,
      document_id: document.document_id, revision: document.revision,
      source_parts: { main_part: document.source.main_part, numbering_part: numberingPart },
      paragraphs: [{ paragraph_id: paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: {}, numbering }],
      runs: [], tables: [], fonts: [], diagnostics: [],
    } as unknown as NativeDocxResolvedLayoutInputV1
    const attest = {
      relationships_part: relationshipsPart, relationships_sha256: digest('1'),
      relationship_id: 'rIdNumbering', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
      relationship_target: 'numbering.xml', part_name: numberingPart,
      content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml' as const, part_sha256: digest('2'),
    }
    numbering.definition_sha256 = nativeDocxResolvedNumberingDefinitionSha256V1(resolved.paragraphs[0]!.numbering!, attest.part_sha256)
    resolved.numbering_source = { ...attest, model_sha256: nativeDocxResolvedNumberingModelSha256V1(resolved.paragraphs, attest) }
    const marker = { ...fact, scope_id: paragraph.id, part_name: paragraph.anchor.part_name, path: paragraph.anchor.path, package_sha256: source.source.package_sha256, source_family: markerFamily }
    return { document, resolved, marker }
  }

  it('repoints only the evidenced marker, re-attests the model, and is reversible', () => {
    const { document, resolved, marker } = model()
    expect(decodeNativeDocxResolvedLayout(resolved).ok).toBe(true)
    const projected = projectNativeDocxEnclosedMarkerFontsV1(document, resolved, [marker], () => true)
    expect(projected.applied).toEqual([{ ...marker, chosen_family: 'Cambria' }])
    expect(projected.resolved.paragraphs[0]!.numbering!.marker_properties.font_family).toBe('Cambria')
    // Nothing else moves: the paragraph mark and the runs keep what they had.
    expect(projected.resolved.paragraphs[0]!.paragraph_mark_properties.font_family).toBeUndefined()
    expect(resolved.paragraphs[0]!.numbering!.marker_properties.font_family).toBe('Liberation Serif')
    expect(projected.resolved.numbering_source!.model_sha256).not.toBe(resolved.numbering_source!.model_sha256)
    // The strict view restores the authored family and the original attestation,
    // so the strict font inventory still exact-joins its scopes.
    expect(stripNativeDocxEnclosedMarkerFontsV1(projected.resolved, [marker])).toEqual(resolved)
    expect(stripNativeDocxEnclosedMarkerFontsV1(resolved, [])).toBe(resolved)
  })

  it('leaves the authored face alone when the host manifest does not attest the substitute', () => {
    const { document, resolved, marker } = model()
    const projected = projectNativeDocxEnclosedMarkerFontsV1(document, resolved, [marker], () => false)
    expect(projected.applied).toEqual([])
    expect(projected.resolved).toEqual(resolved)
    expect(projectNativeDocxEnclosedMarkerFontsV1(document, resolved, [], () => true).applied).toEqual([])
  })

  it('refuses evidence that does not join the resolved marker or its anchor', () => {
    const { document, resolved, marker } = model()
    expect(() => projectNativeDocxEnclosedMarkerFontsV1(document, resolved, [{ ...marker, source_family: 'Calibri' }], () => true))
      .toThrow('does not join the resolved marker family')
    expect(() => projectNativeDocxEnclosedMarkerFontsV1(document, resolved, [{ ...marker, path: '/w:document[1]/w:body[1]/w:p[9]' }], () => true))
      .toThrow('scope anchor does not exact-join')
    expect(() => projectNativeDocxEnclosedMarkerFontsV1(document, resolved, [{ ...marker, package_sha256: hash }], () => true))
      .toThrow('does not exact-join')
    // A marker already carrying the substitute is not evidence for repointing it again.
    const already = model('Cambria')
    expect(() => projectNativeDocxEnclosedMarkerFontsV1(already.document, already.resolved, [{ ...already.marker, source_family: 'Liberation Serif' }], () => true))
      .toThrow('does not join the resolved marker family')
  })
})
