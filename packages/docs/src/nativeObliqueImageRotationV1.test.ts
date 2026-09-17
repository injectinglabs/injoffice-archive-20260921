import { describe, expect, it } from 'vitest'
import { DOCX_NATIVE_PROTOCOL, DOCX_NATIVE_VERSION, decodeNativeDocxDocument, type NativeDocxDocumentV1 } from './nativeContract.js'
import { qualifyNativeDocxInlineImageV1 } from './nativeImagePagePaintV1.js'
import { resolveNativeDocxFloatingAnchorV1 } from './nativeFloatingAnchorV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const RELATIONSHIPS_HASH = `sha256:${'b'.repeat(64)}`
const MEDIA_HASH = `sha256:${'c'.repeat(64)}`
const REVISION = `rev:${'a'.repeat(32)}`
const READ_ONLY = { mode: 'read-only' as const, allowed_operations: [], refusal: { code: 'EXTRACT_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } }

function anchor(path: string, start: number, end: number) {
  return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH }
}

/** crop-pixel.docx in miniature: one floating picture turned 10.685 degrees,
 * wrapped on both sides, with Word's own rotation envelope in wp:effectExtent. */
function documentWith(drawing: Record<string, unknown>): NativeDocxDocumentV1 {
  const run = {
    kind: 'drawing' as const, id: 'run:image', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 105, 109),
    drawing: {
      id: 'drawing:1', anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]', 106, 108),
      relationship_id: 'rImage', media_part: 'word/media/image.png', content_type: 'image/png', placement: 'floating' as const,
      width_emu: 914_400, height_emu: 457_200, x_emu: 914_400, y_emu: 1_270_000,
      horizontal_relative_from: 'column', vertical_relative_from: 'paragraph',
      wrap: 'square', wrap_distance_left_emu: 114_300, wrap_distance_right_emu: 114_300,
      floating_layer: 'front', stacking_order: 7,
      rotation_60000ths: 641_099, floating_effect_extent_emu: { left: 127_000, top: 88_900, right: 127_000, bottom: 88_900 },
      edit_policy: READ_ONLY, ...drawing,
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

function qualify(drawing: Record<string, unknown>) {
  const document = documentWith(drawing)
  return qualifyNativeDocxInlineImageV1(document, 'run:image', onlyDrawing(document))
}

describe('oblique floating picture rotation', () => {
  it('carries the source angle into the paint transform without rounding it to a degree', () => {
    const qualified = qualify({})
    expect(qualified.ok).toBe(true)
    if (!qualified.ok) return
    expect(qualified.value.transform).toEqual({ rotation_degrees: 0, rotation_60000ths: 641_099, flip_horizontal: false, flip_vertical: false })
    // The painted box is still wp:extent; the envelope never resizes it.
    expect(qualified.value.width_millipoints).toBe(72_000)
    expect(qualified.value.height_millipoints).toBe(36_000)
  })

  it('projects the rotation envelope as milli-points on the floating anchor', () => {
    const qualified = qualify({})
    expect(qualified.ok).toBe(true)
    if (!qualified.ok || !qualified.value.floating) throw new Error('expected a floating anchor')
    expect(qualified.value.floating).toMatchObject({
      effect_extent_left_millipoints: 10_000, effect_extent_top_millipoints: 7_000,
      effect_extent_right_millipoints: 10_000, effect_extent_bottom_millipoints: 7_000,
    })
  })

  it('widens the wrap region by the rotation envelope and then by distL/distR', () => {
    const qualified = qualify({})
    if (!qualified.ok || !qualified.value.floating) throw new Error('expected a floating anchor')
    const page = { id: 'page:1', columns: [{ id: 'column:1' }], body_box: { x_millipoints: 72_000 } } as never
    const resolved = resolveNativeDocxFloatingAnchorV1(qualified.value.floating, qualified.value.width_millipoints, page, { page_id: 'page:1', paragraph_top_millipoints: 72_000 })
    // x = body 72000 + offset 72000; the band runs from x - envelope - distL to
    // x + width + envelope + distR.
    expect(resolved.x_millipoints).toBe(144_000)
    expect(resolved.exclusion_left_millipoints).toBe(144_000 - 10_000 - 9_000)
    expect(resolved.exclusion_right_millipoints).toBe(144_000 + 72_000 + 10_000 + 9_000)
  })

  it('refuses square wrapping around an oblique rotation with no source envelope', () => {
    expect(qualify({ floating_effect_extent_emu: undefined })).toMatchObject({ ok: false, code: 'unsupported-image' })
  })

  it('still refuses square wrapping around a quarter turn, whose extent is already rotated', () => {
    expect(qualify({ rotation_degrees: 90, rotation_60000ths: 5_400_000 })).toMatchObject({ ok: false, code: 'unsupported-image' })
  })

  it('paints an oblique rotation that wraps nothing', () => {
    const qualified = qualify({ wrap: 'none', wrap_distance_left_emu: undefined, wrap_distance_right_emu: undefined, floating_effect_extent_emu: undefined })
    expect(qualified.ok).toBe(true)
    if (!qualified.ok) return
    expect(qualified.value.transform.rotation_60000ths).toBe(641_099)
  })

  it('refuses an angle that is not a positive fixed angle below one full turn', () => {
    for (const angle of [-1, 21_600_000, 641_099.5]) expect(qualify({ rotation_60000ths: angle })).toMatchObject({ ok: false, code: 'unsupported-image' })
  })

  it('refuses a quarter-turn angle that disagrees with its whole-degree projection', () => {
    expect(qualify({ rotation_60000ths: 5_400_000 })).toMatchObject({ ok: false, code: 'unsupported-image' })
    expect(qualify({ rotation_degrees: 180, rotation_60000ths: 5_400_000 })).toMatchObject({ ok: false, code: 'unsupported-image' })
  })

  it('validates the contract fields it introduces', () => {
    expect(decodeNativeDocxDocument(documentWith({}) as unknown).ok).toBe(true)
    for (const invalid of [
      { rotation_60000ths: 21_600_000 },
      { rotation_60000ths: 5_400_000 },
      { rotation_degrees: 90, rotation_60000ths: 641_099 },
    ]) expect(decodeNativeDocxDocument(documentWith(invalid) as unknown).ok).toBe(false)
  })

  it('refuses a floating rotation envelope on an inline drawing', () => {
    expect(decodeNativeDocxDocument(documentWith({
      placement: 'inline', x_emu: undefined, y_emu: undefined, wrap: undefined,
      horizontal_relative_from: undefined, vertical_relative_from: undefined,
      wrap_distance_left_emu: undefined, wrap_distance_right_emu: undefined,
      floating_layer: undefined, stacking_order: undefined,
    }) as unknown).ok).toBe(false)
  })
})

describe('square-wrap plan derivation on a refused layout', () => {
  it('leaves a refused pagination to report its refusal instead of deriving a plan from it', async () => {
    const { deriveNativeSquareWrapPlanV1 } = await import('./nativeSquareWrapV1.js')
    // The compiler and the paint-request validator both skip this call unless
    // the layout paginated; the function itself is allowed to be strict.
    expect(() => deriveNativeSquareWrapPlanV1(documentWith({}), { paragraphs: [] } as never, { paragraphs: [] } as never, { status: 'refused' } as never))
      .toThrow('Square wrapping requires complete provisional pagination')
  })
})
