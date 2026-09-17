import { describe, expect, it } from 'vitest'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, type NativeDocxDocumentV1 } from './nativeContract.js'
import { qualifyNativeDocxInlineImageV1 } from './nativeImagePagePaintV1.js'
import {
  DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING,
  nearestNativeDocxMilliPointEmuV1,
  projectNativeDocxApproximateImageExtentsV1,
  validNativeDocxApproximatedImageExtentsV1,
} from './nativeApproximateImageExtentV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const RELATIONSHIPS_HASH = `sha256:${'b'.repeat(64)}`
const MEDIA_HASH = `sha256:${'c'.repeat(64)}`
const REVISION = `rev:${'a'.repeat(32)}`
const READ_ONLY = { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } }

function anchor(path: string, start: number, end: number) {
  return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH }
}

/** One body paragraph holding one inline picture with the given EMU extent. */
function documentWith(drawing: Record<string, unknown>): NativeDocxDocumentV1 {
  const run = {
    kind: 'drawing' as const, id: 'run:image', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 105, 109),
    drawing: {
      id: 'drawing:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 106, 108),
      relationship_id: 'rImage', media_part: 'word/media/image.png', content_type: 'image/png', placement: 'inline' as const,
      width_emu: 2_751_151, height_emu: 2_063_363, edit_policy: READ_ONLY, ...drawing,
    },
  }
  for (const [key, value] of Object.entries(drawing)) if (value === undefined) delete (run.drawing as Record<string, unknown>)[key]
  const paragraph = { id: 'paragraph:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]', 100, 190), edit_policy: READ_ONLY, properties: {}, runs: [run] }
  return {
    protocol: DOCX_NATIVE_PROTOCOL, version: DOCX_NATIVE_VERSION,
    document_id: 'document:test', revision: REVISION,
    source: { package_sha256: HASH, main_part: 'word/document.xml' },
    body: { id: 'story:body', kind: 'body', part_name: 'word/document.xml', anchor: anchor('/w:document[1]/w:body[1]', 1, 3_000), blocks: [{ kind: 'paragraph', id: paragraph.id, paragraph }] },
    sections: [{
      id: 'section:1', anchor: anchor('/w:document[1]/w:body[1]/w:sectPr[1]', 2_100, 2_190), starts_at_block_id: paragraph.id, break_type: 'next-page', title_page: false,
      page: {
        width_twips: 12_240, height_twips: 15_840, orientation: 'portrait',
        margins: { top_twips: 1_440, right_twips: 1_440, bottom_twips: 1_440, left_twips: 1_440, header_twips: 720, footer_twips: 720, gutter_twips: 0 },
        columns: 1, column_spacing_twips: 720, column_layout: 'equal-width', column_definitions: [{ id: 'column:section:1:0', ordinal: 0 }],
      },
      header_refs: [], footer_refs: [],
    }],
    headers: [], footers: [], notes: [], comment_stories: [], comments: [], capabilities: [],
    passthrough_parts: [
      { part_name: 'word/media/image.png', content_type: 'image/png', byte_length: 68, sha256: MEDIA_HASH, policy: 'preserve-verbatim' },
      { part_name: 'word/_rels/document.xml.rels', content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 200, sha256: RELATIONSHIPS_HASH, policy: 'preserve-verbatim' },
    ],
    unsupported: [],
  } as unknown as NativeDocxDocumentV1
}

function onlyDrawing(document: NativeDocxDocumentV1) {
  return document.body.blocks[0]!.paragraph!.runs[0]!.drawing!
}

describe('approximate picture extent rounding', () => {
  it('moves an off-lattice extent to the nearest milli-point, records both EMU values, and leaves strict refusing', () => {
    // FigureAsLabelPicture.docx: cx=2751151 cy=2063363, neither a multiple of 127.
    const source = documentWith({})
    expect(qualifyNativeDocxInlineImageV1(source, 'run:image', onlyDrawing(source))).toMatchObject({ ok: false, code: 'unsupported-image' })
    const projected = projectNativeDocxApproximateImageExtentsV1(source)
    expect(projected.applied).toEqual([
      { run_id: 'run:image', drawing_id: 'drawing:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', field: 'width_emu', source_emu: 2_751_151, painted_emu: 2_751_201 },
      { run_id: 'run:image', drawing_id: 'drawing:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', field: 'height_emu', source_emu: 2_063_363, painted_emu: 2_063_369 },
    ])
    // Each edge moves by well under half a milli-point of EMU: 0.005 pt.
    for (const fact of projected.applied) expect(Math.abs(fact.painted_emu - fact.source_emu)).toBeLessThanOrEqual(63)
    const qualified = qualifyNativeDocxInlineImageV1(projected.document, 'run:image', onlyDrawing(projected.document))
    if (!qualified.ok) throw new Error(qualified.message)
    expect(qualified.value.width_millipoints).toBe(216_630)
    expect(qualified.value.height_millipoints).toBe(162_470)
    // The strict lane's own input is untouched and still refuses it.
    expect(qualifyNativeDocxInlineImageV1(source, 'run:image', onlyDrawing(source)).ok).toBe(false)
    expect(onlyDrawing(source).width_emu).toBe(2_751_151)
    expect(validNativeDocxApproximatedImageExtentsV1(projected.applied)).toBe(true)
  })

  it('never changes which picture is painted: the per-asset exact join is identical', () => {
    const source = documentWith({ width_emu: 127_000, height_emu: 127_000 })
    const exact = qualifyNativeDocxInlineImageV1(source, 'run:image', onlyDrawing(source))
    const rounded = documentWith({})
    const qualified = qualifyNativeDocxInlineImageV1(projectNativeDocxApproximateImageExtentsV1(rounded).document, 'run:image', onlyDrawing(projectNativeDocxApproximateImageExtentsV1(rounded).document))
    if (!exact.ok || !qualified.ok) throw new Error('Expected both pictures to qualify')
    for (const key of ['asset_id', 'part_name', 'relationship_id', 'relationship_part', 'relationship_sha256', 'content_type', 'content_digest', 'byte_length'] as const) {
      expect(qualified.value[key]).toEqual(exact.value[key])
    }
  })

  it('leaves an already exact extent, and every document without a movable picture, byte-identical', () => {
    const exact = documentWith({ width_emu: 127_000, height_emu: 254_000 })
    const projected = projectNativeDocxApproximateImageExtentsV1(exact)
    expect(projected.applied).toEqual([])
    expect(projected.document).toEqual(exact)
  })

  it.each([
    ['an extent that would collapse to zero', { width_emu: 63 }],
    ['an unsupported media type', { content_type: 'image/gif' }],
    ['a drawing with no embedded relationship identity', { media_part: undefined, relationship_id: undefined, content_type: undefined }],
    ['an extent beyond the geometry bound', { width_emu: 12_700_001_000 }],
  ])('still refuses %s rather than rounding it into a paint', (_label, patch) => {
    const source = documentWith(patch)
    expect(qualifyNativeDocxInlineImageV1(source, 'run:image', onlyDrawing(source)).ok).toBe(false)
    const projected = projectNativeDocxApproximateImageExtentsV1(source)
    expect(projected.applied).toEqual([])
    expect(projected.document).toEqual(source)
    expect(qualifyNativeDocxInlineImageV1(projected.document, 'run:image', onlyDrawing(projected.document)).ok).toBe(false)
  })

  it.each([
    ['a negative extent', { width_emu: -2_751_151 }],
    ['a fractional extent', { width_emu: 2_751_151.5 }],
    ['a rotation outside the quarter turns', { rotation_degrees: 45 }],
  ])('never sees %s: the source contract refuses it before this lane', (_label, patch) => {
    expect(() => projectNativeDocxApproximateImageExtentsV1(documentWith(patch))).toThrow(/valid native document/)
  })

  it('rounds a floating picture offset and its effect extents only when the whole drawing then qualifies', () => {
    const floating = documentWith({
      placement: 'floating', width_emu: 650_296, height_emu: 650_296, x_emu: 914_401, y_emu: 914_401,
      horizontal_relative_from: 'page', vertical_relative_from: 'page', wrap: 'none', floating_layer: 'front', stacking_order: 0,
      inline_effect_extent_emu: undefined,
    })
    const projected = projectNativeDocxApproximateImageExtentsV1(floating)
    expect(projected.applied.map(fact => fact.field)).toEqual(['width_emu', 'height_emu', 'x_emu', 'y_emu'])
    const qualified = qualifyNativeDocxInlineImageV1(projected.document, 'run:image', onlyDrawing(projected.document))
    if (!qualified.ok) throw new Error(qualified.message)
    expect(qualified.value.floating).toMatchObject({ offset_x_millipoints: 72_000, offset_y_millipoints: 72_000, horizontal_origin: 'page', vertical_origin: 'page' })
    // A floating picture that stays unqualified for a source reason of its own
    // keeps its authored geometry and its refusal.
    const unanchored = documentWith({ ...{ placement: 'floating', width_emu: 650_296, height_emu: 650_296, x_emu: 914_401, y_emu: 914_401, wrap: 'none', floating_layer: 'front', stacking_order: 0 }, horizontal_relative_from: 'leftMargin', vertical_relative_from: 'page' })
    const refused = projectNativeDocxApproximateImageExtentsV1(unanchored)
    expect(refused.applied).toEqual([])
    expect(refused.document).toEqual(unanchored)
  })

  it('rounds inline effect extents alongside the picture box', () => {
    const source = documentWith({ inline_effect_extent_emu: { left: 0, top: 9_526, right: 9_526, bottom: 0 } })
    const projected = projectNativeDocxApproximateImageExtentsV1(source)
    expect(projected.applied.map(fact => fact.field)).toEqual(['width_emu', 'height_emu', 'effect_extent_top', 'effect_extent_right'])
    expect(onlyDrawing(projected.document).inline_effect_extent_emu).toEqual({ left: 0, top: 9_525, right: 9_525, bottom: 0 })
  })

  it('accepts only facts that are exactly one nearest-lattice move', () => {
    const [fact] = projectNativeDocxApproximateImageExtentsV1(documentWith({})).applied
    if (!fact) throw new Error('Expected a rounding fact')
    expect(validNativeDocxApproximatedImageExtentsV1([fact])).toBe(true)
    // Not the nearest value, an already exact source, a foreign field, and a
    // repeated edge are all rejected: the fact is checkable, not a free label.
    expect(validNativeDocxApproximatedImageExtentsV1([{ ...fact, painted_emu: fact.painted_emu + 127 }])).toBe(false)
    expect(validNativeDocxApproximatedImageExtentsV1([{ ...fact, source_emu: 127_000, painted_emu: 127_000 }])).toBe(false)
    expect(validNativeDocxApproximatedImageExtentsV1([{ ...fact, field: 'rotation_degrees' }])).toBe(false)
    expect(validNativeDocxApproximatedImageExtentsV1([fact, fact])).toBe(false)
    expect(validNativeDocxApproximatedImageExtentsV1([{ ...fact, extra: 1 }])).toBe(false)
  })

  it('states the rounding bound it applies', () => {
    expect(DOCX_APPROXIMATE_IMAGE_EXTENT_WARNING).toContain('nearest milli-point')
    expect(nearestNativeDocxMilliPointEmuV1(650_296)).toBe(650_240)
    expect(nearestNativeDocxMilliPointEmuV1(127_000)).toBeUndefined()
    expect(nearestNativeDocxMilliPointEmuV1(0)).toBeUndefined()
    expect(nearestNativeDocxMilliPointEmuV1(-127_001)).toBeUndefined()
  })
})
