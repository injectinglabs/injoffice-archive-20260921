import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import { createHarfBuzzTextShaperV1 } from '@injoffice/font-metrics/harfbuzz'
import {
  NATIVE_TEXT_LAYOUT_VERSION,
  type FontResource,
  type NativeFontManifest,
  type NativeFontResolver,
  type NativeTextShaper,
  type ResolvedFontFace,
  type ShapedSegment,
} from '@injoffice/font-metrics/layout'
import { type NativeDocxDocumentV1, type NativeDocxRunV1 } from './nativeContract.js'
import {
  DOCX_SHAPING_REQUEST_PROTOCOL,
  shapeNativeDocxLinesV1,
  type NativeDocxShapingRequestV1,
} from './nativeShapingLines.js'
import { decodeNativeDocxShapedLines } from './nativeShapedLinesContract.js'
import {
  DOCX_RESOLVED_LAYOUT_PROTOCOL,
  decodeNativeDocxResolvedLayout,
  nativeDocxResolvedNumberingDefinitionSha256V1,
  nativeDocxResolvedNumberingModelSha256V1,
  type NativeDocxResolvedLayoutInputV1,
  type NativeDocxResolvedNumberingV1,
  type NativeDocxResolvedRunPropertiesV1,
} from './nativeResolvedLayout.js'

const fixture = JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json', import.meta.url), 'utf8')) as NativeDocxDocumentV1
const require = createRequire(import.meta.url)
const digest = 'sha256:054edec1d0211f624fed0cbca9d4f9400b0e491c43742af2c5b0abebf0c990d8' as const
const relationshipsDigest = 'sha256:1111111111111111111111111111111111111111111111111111111111111111'
const numberingDigest = 'sha256:2222222222222222222222222222222222222222222222222222222222222222'
const definitionDigest = 'sha256:3333333333333333333333333333333333333333333333333333333333333333'

const face: ResolvedFontFace = {
  faceId: 'carlito.regular',
  family: 'Carlito',
  weight: 400,
  style: 'normal',
  stretch: 100,
  sourceKind: 'bundled',
  resourceId: 'fonts/carlito.ttf',
  contentDigest: digest,
  resolution: 'exact',
  matchedFamily: 'Carlito',
}

const manifest: NativeFontManifest = {
  version: NATIVE_TEXT_LAYOUT_VERSION,
  manifestId: 'fixture.manifest',
  revision: '1',
  faces: [{ faceId: face.faceId, family: face.family, weight: face.weight, style: face.style, stretch: face.stretch, source: { kind: 'bundled', resourceId: face.resourceId, contentDigest: digest } }],
  fallbackChains: [],
}

function nativeDocument(text = 'ab cd'): NativeDocxDocumentV1 {
  const document = structuredClone(fixture)
  const paragraph = document.body.blocks[0]!.paragraph!
  const textRun = paragraph.runs[0]!
  textRun.text = text
  const lineBreak: NativeDocxRunV1 = {
    kind: 'control',
    id: 'run:intro:break',
    anchor: structuredClone(paragraph.runs[1]!.anchor),
    control: 'line-break',
  }
  const trailing: NativeDocxRunV1 = {
    kind: 'text',
    id: 'run:intro:trailing',
    anchor: structuredClone(paragraph.runs[2]!.anchor),
    text: 'ef',
  }
  paragraph.runs = [textRun, lineBreak, trailing]
  document.body.blocks = [document.body.blocks[0]!]
  document.headers = []
  document.footers = []
  document.notes = []
  document.comment_stories = []
  document.comments = []
  document.sections[0]!.header_refs = []
  document.sections[0]!.footer_refs = []
  document.passthrough_parts = []
  document.unsupported = []
  return document
}

function properties(overrides: Partial<NativeDocxResolvedRunPropertiesV1> = {}): NativeDocxResolvedRunPropertiesV1 {
  return { font_family: 'Carlito', font_size_half_points: 20, language: 'en-US', ...overrides }
}

function resolvedNumbering(paragraphID: string, counterValue: number, resolvedText: string, suffix: NativeDocxResolvedNumberingV1['suffix'] = 'tab'): NativeDocxResolvedNumberingV1 {
  return {
    marker_id: `marker:${paragraphID}`,
    definition_sha256: definitionDigest,
    num_id: '7', abstract_num_id: '3', level: 0, start: 1,
    format: 'decimal', text: '%1.', suffix, alignment: 'start', never_restart: true,
    counter_value: counterValue, counter_values: [{ level: 0, value: counterValue, format: 'decimal' }], resolved_text: resolvedText,
    label_start_twips: 0, label_end_twips: 720, text_start_twips: 720,
    marker_properties: properties({ bold: true }),
  }
}

function attestNumbering(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1): void {
  const relationshipsPart = 'word/_rels/document.xml.rels'
  const numberingPart = 'word/numbering.xml'
  document.passthrough_parts = [
    { part_name: relationshipsPart, content_type: 'application/vnd.openxmlformats-package.relationships+xml', byte_length: 1, sha256: relationshipsDigest, policy: 'preserve-verbatim' },
    { part_name: numberingPart, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', byte_length: 1, sha256: numberingDigest, policy: 'preserve-verbatim' },
  ]
  resolved.source_parts.numbering_part = numberingPart
  const source = {
    relationships_part: relationshipsPart, relationships_sha256: relationshipsDigest,
    relationship_id: 'rIdNumbering', relationship_type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', relationship_target: 'numbering.xml',
    part_name: numberingPart, content_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml' as const, part_sha256: numberingDigest,
  }
  for (const paragraph of resolved.paragraphs) if (paragraph.numbering) paragraph.numbering.definition_sha256 = nativeDocxResolvedNumberingDefinitionSha256V1(paragraph.numbering, source.part_sha256)
  resolved.numbering_source = { ...source, model_sha256: nativeDocxResolvedNumberingModelSha256V1(resolved.paragraphs, source) }
}

function resolvedLayout(document: NativeDocxDocumentV1): NativeDocxResolvedLayoutInputV1 {
  const paragraph = document.body.blocks[0]!.paragraph!
  return {
    protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL,
    version: 1,
    document_id: document.document_id,
    revision: document.revision,
    source_parts: { main_part: document.source.main_part },
    paragraphs: [{
      paragraph_id: paragraph.id,
      applied_styles: ['Normal'],
      properties: {
        alignment: 'center',
        spacing_before_twips: 20,
        spacing_after_twips: 10,
        line: 240,
        line_rule: 'auto',
        indent_left_twips: 10,
        indent_right_twips: 10,
        first_line_twips: 2,
      },
      paragraph_mark_properties: properties(),
    }],
    runs: paragraph.runs.map((run) => ({
      run_id: run.id,
      paragraph_id: paragraph.id,
      applied_paragraph_styles: ['Normal'],
      applied_character_styles: [],
      properties: properties(),
    })),
    tables: [],
    fonts: [{ name: 'Carlito', alt_name: 'Calibri' }],
    diagnostics: [],
  }
}

function makeTextOnly(document: NativeDocxDocumentV1, text: string): void {
  const paragraph = document.body.blocks[0]!.paragraph!
  paragraph.runs = [paragraph.runs[0]!]
  paragraph.runs[0]!.text = text
}

describe('authored DOCX kerning threshold', () => {
  it('rejects malformed thresholds for runs and paragraph marks', () => {
    for(const value of [0,-1,3277,1.5,'20',null]) {
      for(const target of ['run','mark']) {
        const resolved=resolvedLayout(nativeDocument())
        const props=target==='run'?resolved.runs[0]!.properties:resolved.paragraphs[0]!.paragraph_mark_properties
        Object.assign(props,{kerning_min_size_half_points:value})
        expect(decodeNativeDocxResolvedLayout(resolved).ok).toBe(false)
      }
    }
  })
  it('explicitly disables absent/below-threshold kerning and enables inclusive thresholds', async () => {
    for (const [threshold, expected] of [[undefined,0],[21,0],[20,1],[1,1]] as const) {
      const document = nativeDocument()
      makeTextOnly(document,'AV')
      const resolved = resolvedLayout(document)
      resolved.runs[0]!.properties.kerning_min_size_half_points = threshold
      const providers = fakeProviders([])
      const seen: number[] = []
      const originalShape = providers.shaper.shape.bind(providers.shaper)
      providers.shaper.shape = input => { if(input.run.text==='AV')seen.push(input.run.features?.find(feature=>feature.tag==='kern')?.value ?? -1); return originalShape(input) }
      const result = await shapeNativeDocxLinesV1(request(document,resolved),providers)
      expect(result.ok).toBe(true)
      expect(seen.length).toBeGreaterThan(0)
      expect(seen.every(value=>value===expected)).toBe(true)
    }
  })
})

function appendParagraph(document: NativeDocxDocumentV1, id: string, runID: string, text: string): void {
  const source = document.body.blocks[0]!.paragraph!
  const paragraph = structuredClone(source)
  paragraph.id = id
  paragraph.anchor.path = `/w:document[1]/w:body[1]/w:p[${document.body.blocks.length + 1}]`
  paragraph.runs = [paragraph.runs[0]!]
  paragraph.runs[0]!.id = runID
  paragraph.runs[0]!.text = text
  paragraph.runs[0]!.anchor.path = `${paragraph.anchor.path}/w:r[1]`
  document.body.blocks.push({ kind: 'paragraph', id, paragraph })
}

function fakeProviders(requests: Array<{ text: string; family: string; size: number; script: string; direction: string }>): { resolver: NativeFontResolver; shaper: NativeTextShaper } {
  const resource: FontResource = {
    face,
    bytes: new Uint8Array([0, 1, 2, 3]),
    metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 },
  }
  return {
    resolver: {
      providerId: 'fixture-resolver',
      providerRevision: '1',
      resolve: ({ run }) => {
        requests.push({ text: run.text, family: run.font.families[0]!, size: run.fontSizeMilliPoints, script: run.script, direction: run.direction })
        return { status: 'resolved', face, attemptedFaceIds: [face.faceId], decisions: [] }
      },
      load: () => resource,
    },
    shaper: {
      providerId: 'fixture-shaper',
      providerRevision: '1',
      shape: ({ run }): ShapedSegment => {
        const clusters = [...run.text].map((character, index, characters) => {
          const startUtf16 = characters.slice(0, index).join('').length
          const endUtf16 = startUtf16 + character.length
          return { startUtf16, endUtf16, glyphStart: index, glyphEnd: index + 1, advanceInlineMilliPoints: 1_000, whitespace: /^\s$/u.test(character) }
        })
        return {
          startUtf16: 0,
          endUtf16: run.text.length,
          face,
          glyphs: clusters.map((_cluster, index) => ({ glyphId: index + 1, clusterIndex: index, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 })),
          clusters,
          metrics: { fontSizeMilliPoints: run.fontSizeMilliPoints, ascentMilliPoints: 8_000, descentMilliPoints: -2_000, lineGapMilliPoints: 2_000, lineHeightMilliPoints: 12_000 },
          advanceInlineMilliPoints: clusters.length * 1_000,
          advanceBlockMilliPoints: 0,
        }
      },
    },
  }
}

function request(document: NativeDocxDocumentV1, resolved = resolvedLayout(document)): NativeDocxShapingRequestV1 {
  return {
    protocol: DOCX_SHAPING_REQUEST_PROTOCOL,
    version: 1,
    document,
    resolved_layout: resolved,
    font_manifest: manifest,
    available_width_millipoints: 5_000,
    tab_interval_millipoints: 2_000,
  }
}

function unreadRecursiveProviderBomb(onTraversal: () => void): Record<string, unknown> {
  const bomb: Record<string, unknown> = {}
  bomb.self = bomb
  Object.defineProperty(bomb, 'huge', {
    enumerable: true,
    get() {
      onTraversal()
      throw new Error('rejected provider data was traversed')
    },
  })
  return bomb
}

describe('native DOCX resolved-layout projection', () => {
  it('orders Unicode validation issues by UTF-16 code units, independent of process locale', async () => {
    const document = nativeDocument()
    const value = request(document) as NativeDocxShapingRequestV1 & Record<string, unknown>
    value['Å'] = true
    value.Z = true
    value['😀'] = true
    value.A = true
    const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.filter((issue) => issue.code === 'UNKNOWN_FIELD').map((issue) => issue.path)).toEqual(['/A', '/Z', '/Å', '/😀'])
  })

  it('rejects unknown keys and dangling run-to-paragraph references', () => {
    const document = nativeDocument()
    const value = resolvedLayout(document) as unknown as Record<string, any>
    value.paragraphs[0].css = 'not-v1'
    value.runs[0].paragraph_id = 'paragraph:missing'
    const decoded = decodeNativeDocxResolvedLayout(value)
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'UNKNOWN_FIELD', path: '/paragraphs/0/css' }),
      expect.objectContaining({ code: 'BROKEN_REFERENCE', path: '/runs/0/paragraph_id' }),
    ]))
  })

  it('rejects null, negative zero, non-canonical parts, and Go-out-of-range numbering ids', () => {
    const document = nativeDocument()
    const value = resolvedLayout(document) as unknown as Record<string, any>
    value.source_parts.styles_part = 'word/%2E%2E/styles.xml'
    value.paragraphs[0].properties.spacing_before_twips = -0
    value.paragraphs[0].numbering = {
      num_id: '2147483648',
      abstract_num_id: '0',
      level: 0,
      marker_properties: {},
    }
    value.fonts[0].alt_name = null
    const decoded = decodeNativeDocxResolvedLayout(value)
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues.map((issue) => issue.path)).toEqual(expect.arrayContaining([
      '/source_parts/styles_part',
      '/paragraphs/0/properties/spacing_before_twips',
      '/paragraphs/0/numbering/num_id',
      '/fonts/0/alt_name',
    ]))
  })

  it('applies the Go UTF-8 byte and Unicode-scalar bounds to numbering text', () => {
    const document = nativeDocument()
    const value = resolvedLayout(document)
    value.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, '1.')
    value.paragraphs[0]!.numbering!.text = '😀'.repeat(300)
    attestNumbering(document, value)
    const decoded = decodeNativeDocxResolvedLayout(value)
    expect(decoded.ok).toBe(false)
    if (!decoded.ok) expect(decoded.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/paragraphs/0/numbering/text' })]))

    const malformed = resolvedLayout(document)
    malformed.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, '\ud800')
    attestNumbering(document, malformed)
    const malformedDecoded = decodeNativeDocxResolvedLayout(malformed)
    expect(malformedDecoded.ok).toBe(false)
    if (!malformedDecoded.ok) expect(malformedDecoded.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/paragraphs/0/numbering/resolved_text' })]))

    const oversized = resolvedLayout(document)
    oversized.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, '1.')
    oversized.paragraphs[0]!.numbering!.text = 'a'.repeat(1025)
    attestNumbering(document, oversized)
    const oversizedDecoded = decodeNativeDocxResolvedLayout(oversized)
    expect(oversizedDecoded.ok).toBe(false)
    if (!oversizedDecoded.ok) expect(oversizedDecoded.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/paragraphs/0/numbering/text' })]))

    const tooManyScalars = resolvedLayout(document)
    tooManyScalars.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, 'a'.repeat(32))
    attestNumbering(document, tooManyScalars)
    const tooManyScalarsDecoded = decodeNativeDocxResolvedLayout(tooManyScalars)
    expect(tooManyScalarsDecoded.ok).toBe(false)
    if (!tooManyScalarsDecoded.ok) expect(tooManyScalarsDecoded.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/paragraphs/0/numbering/resolved_text' })]))
  })

  it('requires and validates authoritative paragraph-mark run properties', () => {
    const document = nativeDocument()
    const missing = resolvedLayout(document) as unknown as Record<string, any>
    delete missing.paragraphs[0].paragraph_mark_properties
    const missingResult = decodeNativeDocxResolvedLayout(missing)
    expect(missingResult.ok).toBe(false)
    if (!missingResult.ok) expect(missingResult.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/paragraphs/0/paragraph_mark_properties' })]))
    const invalid = resolvedLayout(document) as unknown as Record<string, any>
    invalid.paragraphs[0].paragraph_mark_properties.font_size_half_points = 0
    const invalidResult = decodeNativeDocxResolvedLayout(invalid)
    expect(invalidResult.ok).toBe(false)
    if (!invalidResult.ok) expect(invalidResult.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/paragraphs/0/paragraph_mark_properties/font_size_half_points' })]))
  })
})

describe('shapeNativeDocxLinesV1', () => {
  it('accepts exact real HarfBuzz glyphs through the strict shaped-lines contract', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'office A\u0301')
    const resolved = resolvedLayout(document)
    for (const entry of resolved.runs) entry.properties.font_family = 'DejaVu Sans'
    resolved.paragraphs[0]!.paragraph_mark_properties.font_family = 'DejaVu Sans'
    resolved.fonts = [{ name: 'DejaVu Sans' }]
    const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
    const realFace: ResolvedFontFace = {
      faceId: 'fixture.dejavu-sans',
      family: 'DejaVu Sans',
      postscriptName: 'DejaVuSans',
      weight: 400,
      style: 'normal',
      stretch: 100,
      sourceKind: 'bundled',
      resourceId: 'fixture/dejavu-sans-2.37',
      contentDigest: 'sha256:7da195a74c55bef988d0d48f9508bd5d849425c1770dba5d7bfc6ce9ed848954',
      resolution: 'exact',
      matchedFamily: 'DejaVu Sans',
    }
    const realManifest: NativeFontManifest = {
      version: NATIVE_TEXT_LAYOUT_VERSION,
      manifestId: 'fixture.dejavu',
      revision: '2.37',
      faces: [{
        faceId: realFace.faceId,
        family: realFace.family,
        postscriptName: realFace.postscriptName,
        weight: 400,
        style: 'normal',
        stretch: 100,
        source: { kind: 'bundled', resourceId: realFace.resourceId, contentDigest: realFace.contentDigest },
      }],
      fallbackChains: [],
    }
    const resource: FontResource = {
      face: realFace,
      bytes,
      metrics: { unitsPerEm: 2_048, ascender: 1_901, descender: -483, lineGap: 0, underlinePosition: -130, underlineThickness: 90 },
    }
    const resolver: NativeFontResolver = {
      providerId: 'fixture.digest-resolver',
      providerRevision: 'resolver:dejavu-2.37',
      resolve: () => ({ status: 'resolved', face: realFace, attemptedFaceIds: [realFace.faceId], decisions: [] }),
      load: () => resource,
    }
    const shaper = createHarfBuzzTextShaperV1({ sourceRevision: 'git:docs-integration' })
    const input = { ...request(document, resolved), font_manifest: realManifest, available_width_millipoints: 100_000 }
    const result = await shapeNativeDocxLinesV1(input, { resolver, shaper })
    expect(result.ok, result.ok ? '' : JSON.stringify(result.issues)).toBe(true)
    if (!result.ok) return
    expect(decodeNativeDocxShapedLines(result.value).ok).toBe(true)
    expect(result.value.providers).toMatchObject({ shaper_id: shaper.providerId, shaper_revision: shaper.providerRevision })
    const fragments = result.value.paragraphs[0]!.lines.flatMap((line) => line.fragments)
    expect(fragments.flatMap((fragment) => fragment.glyphs).map((glyph) => glyph.glyph_id)).toContain(5_044)
    expect(fragments.every((fragment) => fragment.face_id === realFace.faceId)).toBe(true)
    makeTextOnly(document,'AV')
    const advances: number[] = []
    for (const threshold of [undefined,21,20,1]) {
      resolved.runs[0]!.properties.kerning_min_size_half_points = threshold
      const shaped = await shapeNativeDocxLinesV1(input,{resolver,shaper})
      expect(shaped.ok).toBe(true)
      if(shaped.ok) advances.push(shaped.value.paragraphs[0]!.lines.flatMap(line=>line.fragments).flatMap(fragment=>fragment.glyphs).reduce((sum,glyph)=>sum+glyph.advance_x_millipoints,0))
    }
    expect(advances[0]).toBe(advances[1])
    expect(advances[2]).toBe(advances[3])
    expect(advances[2]).toBeLessThan(advances[0]!)
  })

  it('joins durable ids, uses resolved properties, converts units, wraps clusters, and honors hard breaks', async () => {
    const document = nativeDocument()
    const requests: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(request(document), fakeProviders(requests))
    expect(result.ok, result.ok ? '' : JSON.stringify(result.issues)).toBe(true)
    if (!result.ok) return
    expect(requests).toEqual(expect.arrayContaining([
      { text: 'ab', family: 'Carlito', size: 10_000, script: 'Latn', direction: 'ltr' },
      { text: ' ', family: 'Carlito', size: 10_000, script: 'Zyyy', direction: 'ltr' },
      { text: 'cd', family: 'Carlito', size: 10_000, script: 'Latn', direction: 'ltr' },
    ]))
    const paragraph = result.value.paragraphs[0]!
    expect(paragraph).toMatchObject({
      paragraph_id: 'paragraph:intro',
      spacing_before_millipoints: 1_000,
      spacing_after_millipoints: 500,
      indent_start_millipoints: 500,
      indent_end_millipoints: 500,
      first_line_delta_millipoints: 100,
      block_advance_millipoints: 37_500,
    })
    expect(paragraph.lines).toHaveLength(3)
    expect(paragraph.lines.map((line) => line.fragments.map((fragment) => fragment.text).join(''))).toEqual(['ab ', 'cd', 'ef'])
    expect(paragraph.lines[0]).toMatchObject({ available_width_millipoints: 3_900, advance_inline_millipoints: 3_000, inline_offset_millipoints: 1_050, line_height_millipoints: 12_000 })
    expect(paragraph.lines[1]?.hard_break_after).toEqual({ source_run_id: 'run:intro:break', control: 'line-break' })
    expect(paragraph.lines.flatMap((line) => line.fragments).every((fragment) => fragment.source_id.startsWith('run:'))).toBe(true)
    expect(result.value.font_manifest).toEqual({ manifest_id: manifest.manifestId, revision: manifest.revision })
    expect(result.value.providers).toEqual(expect.objectContaining({ resolver_id: 'fixture-resolver', resolver_revision: '1', shaper_id: 'fixture-shaper', shaper_revision: '1', bidi_id: 'injoffice.bidi-js', bidi_unicode_version: '13.0.0' }))
  })

  it('rejects mismatched durable identities before invoking providers', async () => {
    const document = nativeDocument()
    const resolved = resolvedLayout(document)
    resolved.runs[0]!.run_id = 'run:missing'
    const requests: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders(requests))
    expect(result.ok).toBe(false)
    expect(requests).toEqual([])
    if (result.ok) return
    expect(result.issues.some((issue) => issue.code === 'BROKEN_REFERENCE')).toBe(true)
  })

  it('rejects unstable provider identities before invoking provider code', async () => {
    const document = nativeDocument()
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const providers = fakeProviders(calls)
    Object.defineProperty(providers.resolver, 'providerId', { value: 'invalid provider id' })
    const result = await shapeNativeDocxLinesV1(request(document), providers)
    expect(result.ok).toBe(false)
    expect(calls).toEqual([])
    if (!result.ok) expect(result.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/providers/resolver/provider_id' })]))
  })

  it('does not let providers forge live manifest or provider provenance during calls', async () => {
    const manifestDocument = nativeDocument()
    makeTextOnly(manifestDocument, 'manifest mutation')
    const manifestRequest = request(manifestDocument, resolvedLayout(manifestDocument))
    manifestRequest.font_manifest = structuredClone(manifest)
    const manifestProviders = fakeProviders([])
    const forgedFace = { ...face, faceId: 'forged.face', resourceId: 'fonts/forged.ttf' }
    manifestProviders.resolver.resolve = () => {
      const live = manifestRequest.font_manifest as any
      live.manifestId = 'forged.manifest'
      live.revision = 'forged.revision'
      live.faces.push({ faceId: forgedFace.faceId, family: forgedFace.family, weight: forgedFace.weight, style: forgedFace.style, stretch: forgedFace.stretch, source: { kind: forgedFace.sourceKind, resourceId: forgedFace.resourceId, contentDigest: forgedFace.contentDigest } })
      return { status: 'resolved', face: forgedFace, attemptedFaceIds: [forgedFace.faceId], decisions: [] }
    }
    const manifestResult = await shapeNativeDocxLinesV1(manifestRequest, manifestProviders)
    expect(manifestResult.ok).toBe(true)
    if (!manifestResult.ok) return
    expect(manifestResult.value.paragraphs).toEqual([])
    expect(manifestResult.value.font_manifest).toEqual({ manifest_id: manifest.manifestId, revision: manifest.revision })
    expect(manifestResult.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output' })]))

    const identityDocument = nativeDocument()
    makeTextOnly(identityDocument, 'identity mutation')
    const identityProviders = fakeProviders([])
    const mutableResolver = identityProviders.resolver as unknown as { providerId: string; providerRevision: string }
    identityProviders.resolver.resolve = () => {
      mutableResolver.providerId = 'forged-resolver'
      mutableResolver.providerRevision = 'forged-revision'
      return { status: 'resolved', face, attemptedFaceIds: [face.faceId], decisions: [] }
    }
    const identityResult = await shapeNativeDocxLinesV1(request(identityDocument, resolvedLayout(identityDocument)), identityProviders)
    expect(identityResult.ok).toBe(true)
    if (!identityResult.ok) return
    expect(identityResult.value.paragraphs).toEqual([])
    expect(identityResult.value.providers).toEqual(expect.objectContaining({ resolver_id: 'fixture-resolver', resolver_revision: '1', shaper_id: 'fixture-shaper', shaper_revision: '1', bidi_id: 'injoffice.bidi-js', bidi_unicode_version: '13.0.0' }))
    expect(identityResult.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output', message: expect.stringContaining('snapshotted') })]))
  })

  it('passes owned immutable run and face snapshots across resolver boundaries', async () => {
    const resolveDocument = nativeDocument()
    makeTextOnly(resolveDocument, 'resolver run mutation')
    const resolveProviders = fakeProviders([])
    resolveProviders.resolver.resolve = ({ run }) => {
      ;(run as any).text = 'forged text'
      return { status: 'resolved', face, attemptedFaceIds: [face.faceId], decisions: [] }
    }
    const resolveResult = await shapeNativeDocxLinesV1(request(resolveDocument, resolvedLayout(resolveDocument)), resolveProviders)
    expect(resolveResult.ok).toBe(true)
    if (!resolveResult.ok) return
    expect(resolveResult.value.paragraphs).toEqual([])
    expect(resolveResult.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'provider-failure' })]))

    const loadDocument = nativeDocument()
    makeTextOnly(loadDocument, 'load face mutation')
    const loadProviders = fakeProviders([])
    loadProviders.resolver.load = (resolvedFace) => {
      ;(resolvedFace as any).matchedFamily = 'forged family'
      return { face, bytes: new Uint8Array([0, 1, 2, 3]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 } }
    }
    const loadResult = await shapeNativeDocxLinesV1(request(loadDocument, resolvedLayout(loadDocument)), loadProviders)
    expect(loadResult.ok).toBe(true)
    if (!loadResult.ok) return
    expect(loadResult.value.paragraphs).toEqual([])
    expect(loadResult.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'provider-failure' })]))
  })

  it('snapshots provider metadata arrays and isolates the authoritative font cache from shaper bytes', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'ab')
    const providers = fakeProviders([])
    const liveDecisions: any[] = []
    const liveAttempts = [face.faceId]
    providers.resolver.resolve = () => ({ status: 'resolved', face: { ...face }, decisions: liveDecisions, attemptedFaceIds: liveAttempts })
    const loaded: FontResource = { face: { ...face }, bytes: new Uint8Array([0, 1, 2, 3]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 } }
    providers.resolver.load = (resolvedFace) => {
      liveDecisions.push({ code: 'provider-failure', message: 'forged after resolve snapshot', recoverable: false })
      liveAttempts.push('forged.face')
      expect(resolvedFace).not.toBe(face)
      expect(Object.isFrozen(resolvedFace)).toBe(true)
      return loaded
    }
    const originalShape = providers.shaper.shape.bind(providers.shaper)
    let providerFacingResource: FontResource | undefined
    providers.shaper.shape = (shapeRequest) => {
      providerFacingResource = shapeRequest.font
      shapeRequest.font.bytes[0] = 9
      return originalShape(shapeRequest)
    }
    const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toHaveLength(1)
    expect(result.value.font_manifest).toEqual({ manifest_id: manifest.manifestId, revision: manifest.revision })
    expect(result.value.diagnostics).not.toEqual(expect.arrayContaining([expect.objectContaining({ message: expect.stringContaining('forged after resolve snapshot') })]))
    expect(providerFacingResource).toBeDefined()
    expect(providerFacingResource).not.toBe(loaded)
    expect(loaded.bytes[0]).toBe(0)
  })

  it('captures provider metadata getters exactly once before validation and accounting', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'gettersnapshot')
    const providers = fakeProviders([])
    let decisionsReads = 0
    let attemptsReads = 0
    providers.resolver.resolve = (() => ({
      status: 'resolved',
      face,
      get decisions() {
        decisionsReads += 1
        return decisionsReads === 1 ? [] : [{ code: 'provider-failure', message: 'forged second read', recoverable: false }]
      },
      get attemptedFaceIds() {
        attemptsReads += 1
        return attemptsReads === 1 ? [face.faceId] : ['forged.face']
      },
    })) as NativeFontResolver['resolve']
    const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toHaveLength(1)
    expect(decisionsReads).toBe(1)
    expect(attemptsReads).toBe(1)
    expect(result.value.diagnostics).not.toEqual(expect.arrayContaining([expect.objectContaining({ message: expect.stringContaining('forged second read') })]))
  })

  it('rejects recursive unknown resolver face and decision fields without traversing them', async () => {
    for (const kind of ['face', 'decision'] as const) {
      const document = nativeDocument()
      makeTextOnly(document, `${kind}bomb`)
      const providers = fakeProviders([])
      let traversals = 0
      let loads = 0
      let shapes = 0
      const malformedFace: any = { ...face }
      const malformedDecision: any = { code: 'provider-failure', message: 'bounded', recoverable: false }
      if (kind === 'face') malformedFace.unknown_nested = unreadRecursiveProviderBomb(() => { traversals += 1 })
      else malformedDecision.unknown_nested = unreadRecursiveProviderBomb(() => { traversals += 1 })
      providers.resolver.resolve = () => ({ status: 'resolved', face: malformedFace, decisions: kind === 'decision' ? [malformedDecision] : [], attemptedFaceIds: [face.faceId] })
      providers.resolver.load = () => { loads += 1; throw new Error('load must not run') }
      providers.shaper.shape = () => { shapes += 1; throw new Error('shape must not run') }
      const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs).toEqual([])
      expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output' })]))
      expect(traversals).toBe(0)
      expect(loads).toBe(0)
      expect(shapes).toBe(0)
    }
  })

  it('rejects recursive unknown loaded face and metric fields before shaping', async () => {
    for (const kind of ['face', 'metrics'] as const) {
      const document = nativeDocument()
      makeTextOnly(document, `${kind}loadbomb`)
      const providers = fakeProviders([])
      let traversals = 0
      let shapes = 0
      const loadedFace: any = { ...face }
      const loadedMetrics: any = { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 }
      if (kind === 'face') loadedFace.unknown_nested = unreadRecursiveProviderBomb(() => { traversals += 1 })
      else loadedMetrics.unknown_nested = unreadRecursiveProviderBomb(() => { traversals += 1 })
      providers.resolver.load = () => ({ face: loadedFace, bytes: new Uint8Array([0, 1, 2, 3]), metrics: loadedMetrics })
      providers.shaper.shape = () => { shapes += 1; throw new Error('shape must not run') }
      const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs).toEqual([])
      expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output' })]))
      expect(traversals).toBe(0)
      expect(shapes).toBe(0)
    }
  })

  it('rejects recursive unknown shaped glyph and cluster fields without traversing them', async () => {
    for (const kind of ['glyph', 'cluster'] as const) {
      const document = nativeDocument()
      makeTextOnly(document, `${kind}bomb`)
      const providers = fakeProviders([])
      const originalShape = providers.shaper.shape.bind(providers.shaper)
      let traversals = 0
      providers.shaper.shape = async (shapeRequest) => {
        const shaped = await originalShape(shapeRequest) as ShapedSegment
        const entry: any = kind === 'glyph' ? shaped.glyphs[0] : shaped.clusters[0]
        entry.unknown_nested = unreadRecursiveProviderBomb(() => { traversals += 1 })
        return shaped
      }
      const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs).toEqual([])
      expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output' })]))
      expect(traversals).toBe(0)
    }
  })

  it('refuses a paragraph when the shaper attempts to mutate its owned run snapshot', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'shaper run mutation')
    const providers = fakeProviders([])
    providers.shaper.shape = ({ run }) => {
      ;(run.font.families as string[])[0] = 'Forged'
      throw new Error('unreachable')
    }
    const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'provider-failure' })]))
  })

  it('captures adversarial thrown provider values once and keeps failure diagnostics bounded', async () => {
    let throwingReads = 0
    const throwingMessage = {}
    Object.defineProperty(throwingMessage, 'message', {
      get() {
        throwingReads += 1
        throw new Error('message getter trap')
      },
    })
    let unstableReads = 0
    const unstableMessage = {}
    Object.defineProperty(unstableMessage, 'message', {
      get() {
        unstableReads += 1
        return unstableReads === 1 ? `first:${'x'.repeat(10_000)}` : 'forged-second-read'
      },
    })
    let recursiveProxy: any
    recursiveProxy = new Proxy({}, { get: () => recursiveProxy })

    for (const [thrown, expected] of [[throwingMessage, 'unreadable provider error'], [unstableMessage, 'first:'], [recursiveProxy, 'unknown provider error']] as const) {
      const document = nativeDocument()
      makeTextOnly(document, 'providerthrow')
      const providers = fakeProviders([])
      providers.resolver.resolve = () => { throw thrown }
      const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs).toEqual([])
      const failure = result.value.diagnostics.find((diagnostic) => diagnostic.code === 'provider-failure')
      expect(failure?.message).toContain(expected)
      expect(failure?.message.length).toBeLessThanOrEqual(4_096)
      expect(failure?.message).not.toContain('forged-second-read')
    }
    expect(throwingReads).toBe(1)
    expect(unstableReads).toBe(1)
  })

  it('caches one validated resource per exact face across alternating script spans', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'AاAاAا')
    const resolved = resolvedLayout(document)
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const providers = fakeProviders(calls)
    let loads = 0
    let shapes = 0
    const mutableResource: FontResource = { face: { ...face }, bytes: new Uint8Array([0, 1, 2, 3]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 200 } }
    const shape = providers.shaper.shape.bind(providers.shaper)
    providers.resolver.load = () => { loads += 1; return mutableResource }
    providers.shaper.shape = (providerRequest) => {
      shapes += 1
      const output = shape(providerRequest)
      if (shapes === 1) {
        mutableResource.face = { ...face, matchedFamily: 'mutated-after-load' }
        mutableResource.metrics.ascender = -1
        mutableResource.bytes[0] = 9
      }
      return output
    }
    const result = await shapeNativeDocxLinesV1(request(document, resolved), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(calls).toHaveLength(6)
    expect(loads).toBe(1)
    expect(shapes).toBe(6)
    expect(result.value.paragraphs).toHaveLength(1)
    expect(result.value.diagnostics).not.toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output' })]))
  })

  it('shapes ordinary list markers with deterministic counters and explicit tab stops', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'first')
    appendParagraph(document, 'paragraph:second', 'run:second', 'second')
    const resolved = resolvedLayout(document)
    resolved.paragraphs.push({ paragraph_id: 'paragraph:second', applied_styles: ['Normal'], properties: {}, paragraph_mark_properties: properties() })
    resolved.runs.push({ run_id: 'run:second', paragraph_id: 'paragraph:second', applied_paragraph_styles: ['Normal'], applied_character_styles: [], properties: properties() })
    resolved.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, '1.')
    resolved.paragraphs[1]!.numbering = resolvedNumbering('paragraph:second', 2, '2.')
    for (const paragraph of resolved.paragraphs) paragraph.numbering!.numbering_tab_twips = 840
    for (const paragraph of resolved.paragraphs) paragraph.properties = { indent_start_twips: 720, hanging_twips: 720 }
    attestNumbering(document, resolved)
    const value = request(document, resolved)
    value.available_width_millipoints = 60_000
    value.tab_interval_millipoints = 4_000
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(value, fakeProviders(calls))
    expect(result.ok, result.ok ? '' : JSON.stringify(result.issues)).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs.map((paragraph) => paragraph.lines[0]!.fragments.map((fragment) => fragment.text).join(''))).toEqual(['1.\tfirst', '2.\tsecond'])
    const first = result.value.paragraphs[0]!.lines[0]!
    expect(first.fragments.find((fragment) => fragment.text === '\t')?.advance_inline_millipoints).toBe(40_000)
    expect(first.fragments.slice(0, 2).every((fragment) => fragment.source_kind === 'list-marker' && fragment.source_id === 'paragraph:intro')).toBe(true)
    const alternate = structuredClone(value)
    alternate.tab_interval_millipoints = 9_000
    const alternateResult = await shapeNativeDocxLinesV1(alternate, fakeProviders([]))
    expect(alternateResult.ok).toBe(true)
    if (alternateResult.ok) expect(alternateResult.value.paragraphs[0]!.list_marker!.text_start_millipoints).toBe(result.value.paragraphs[0]!.list_marker!.text_start_millipoints)
    const partialMarker = structuredClone(result.value)
    partialMarker.paragraphs[0]!.list_marker!.text = '12.'
    expect(decodeNativeDocxShapedLines(partialMarker).ok).toBe(false)
  })

  it('atomically refuses explicit numbering tabs at or before the shaped marker end in LTR and RTL', async () => {
    for (const direction of ['ltr', 'rtl'] as const) for (const tabTwips of [39, 40]) {
      const document = nativeDocument()
      makeTextOnly(document, 'body')
      const resolved = resolvedLayout(document)
      resolved.paragraphs[0]!.properties = { indent_start_twips: 720, hanging_twips: 720, ...(direction === 'rtl' ? { bidi: true } : {}) }
      resolved.paragraphs[0]!.numbering = { ...resolvedNumbering('paragraph:intro', 1, '1.'), alignment: 'left', numbering_tab_twips: tabTwips }
      attestNumbering(document, resolved)
      const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
      const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders(calls))
      expect(result.ok, `${direction} ${tabTwips}`).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs, `${direction} ${tabTwips}`).toEqual([])
      expect(result.value.diagnostics, `${direction} ${tabTwips}`).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'list-marker-tab-deferred', scope_id: 'paragraph:intro' })]))
      expect(calls.map((entry) => entry.text).join(''), `${direction} ${tabTwips}`).toBe('1.')
    }
  })

  it('accepts the first explicit numbering-tab boundary strictly after the shaped marker end in LTR and RTL', async () => {
    for (const direction of ['ltr', 'rtl'] as const) {
      const document = nativeDocument()
      makeTextOnly(document, 'body')
      const resolved = resolvedLayout(document)
      resolved.paragraphs[0]!.properties = { indent_start_twips: 720, hanging_twips: 720, ...(direction === 'rtl' ? { bidi: true } : {}) }
      resolved.paragraphs[0]!.numbering = { ...resolvedNumbering('paragraph:intro', 1, '1.'), alignment: 'left', numbering_tab_twips: 41 }
      attestNumbering(document, resolved)
      const value = request(document, resolved)
      value.available_width_millipoints = 60_000
      const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
      expect(result.ok, direction).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs, direction).toHaveLength(1)
      expect(result.value.paragraphs[0]!.list_marker!.text_start_millipoints, direction).toBe(2_050)
      expect(result.value.paragraphs[0]!.lines[0]!.fragments.find((fragment) => fragment.text === '\t')?.advance_inline_millipoints, direction).toBe(50)
    }
  })

  it('binds left, right, and center marker glyphs to the numbering text-margin anchor', async () => {
    for (const [alignment, expectedStart] of [['left', 5_000], ['right', 3_000], ['center', 4_000]] as const) {
      const document = nativeDocument()
      makeTextOnly(document, 'body')
      const resolved = resolvedLayout(document)
      resolved.paragraphs[0]!.properties = { indent_start_twips: 300, hanging_twips: 200 }
      resolved.paragraphs[0]!.numbering = { ...resolvedNumbering('paragraph:intro', 1, '1.', 'nothing'), alignment, label_start_twips: 100, label_end_twips: 300, text_start_twips: 300 }
      attestNumbering(document, resolved)
      const value = request(document, resolved)
      value.available_width_millipoints = 60_000
      const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      const paragraph = result.value.paragraphs[0]!
      const marker = paragraph.list_marker!
      expect(marker.marker_start_millipoints).toBe(expectedStart)
      expect(paragraph.lines[0]!.inline_offset_millipoints).toBe(expectedStart)
      expect(paragraph.lines[0]!.fragments.filter((fragment) => fragment.source_kind === 'list-marker' && fragment.text !== '').map((fragment) => fragment.text).join('')).toBe('1.')
      const shifted = structuredClone(result.value)
      shifted.paragraphs[0]!.lines[0]!.inline_offset_millipoints += 1
      expect(decodeNativeDocxShapedLines(shifted).ok).toBe(false)
    }
  })

  it('rejects numbering model, raw-part, and relationship hash substitutions before shaping', async () => {
    const make = () => {
      const document = nativeDocument()
      makeTextOnly(document, 'bound')
      const resolved = resolvedLayout(document)
      resolved.paragraphs[0]!.properties = { indent_start_twips: 720, hanging_twips: 720 }
      resolved.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, '1.')
      attestNumbering(document, resolved)
      return { document, resolved }
    }
    const model = make()
    model.resolved.numbering_source!.model_sha256 = relationshipsDigest
    expect(decodeNativeDocxResolvedLayout(model.resolved).ok).toBe(false)

    const selfConsistentTargetSubstitution = make()
    selfConsistentTargetSubstitution.resolved.numbering_source!.relationship_target = 'substituted.xml'
    selfConsistentTargetSubstitution.resolved.numbering_source!.model_sha256 = nativeDocxResolvedNumberingModelSha256V1(
      selfConsistentTargetSubstitution.resolved.paragraphs,
      selfConsistentTargetSubstitution.resolved.numbering_source!,
    )
    expect(decodeNativeDocxResolvedLayout(selfConsistentTargetSubstitution.resolved).ok).toBe(false)

    const selfConsistentAlignmentSubstitution = make()
    selfConsistentAlignmentSubstitution.resolved.paragraphs[0]!.numbering!.alignment = 'left'
    selfConsistentAlignmentSubstitution.resolved.numbering_source!.model_sha256 = nativeDocxResolvedNumberingModelSha256V1(selfConsistentAlignmentSubstitution.resolved.paragraphs, selfConsistentAlignmentSubstitution.resolved.numbering_source!)
    expect(decodeNativeDocxResolvedLayout(selfConsistentAlignmentSubstitution.resolved).ok).toBe(false)

    const styleSelectionSubstitution = make()
    styleSelectionSubstitution.resolved.paragraphs[0]!.style_id = 'ListLinked'
    styleSelectionSubstitution.resolved.paragraphs[0]!.applied_styles.push('ListLinked')
    styleSelectionSubstitution.resolved.paragraphs[0]!.numbering!.level_style_id = 'ListLinked'
    attestNumbering(styleSelectionSubstitution.document, styleSelectionSubstitution.resolved)
    expect(decodeNativeDocxResolvedLayout(styleSelectionSubstitution.resolved).ok).toBe(true)
    styleSelectionSubstitution.resolved.paragraphs[0]!.style_id = 'ListForged'
    styleSelectionSubstitution.resolved.paragraphs[0]!.numbering!.level_style_id = 'ListForged'
    styleSelectionSubstitution.resolved.numbering_source!.model_sha256 = nativeDocxResolvedNumberingModelSha256V1(styleSelectionSubstitution.resolved.paragraphs, styleSelectionSubstitution.resolved.numbering_source!)
    expect(decodeNativeDocxResolvedLayout(styleSelectionSubstitution.resolved).ok).toBe(false)

    for (const mutate of [
      ({ document }: ReturnType<typeof make>) => { document.passthrough_parts[1]!.sha256 = relationshipsDigest },
      ({ document }: ReturnType<typeof make>) => { document.passthrough_parts[0]!.sha256 = numberingDigest },
      ({ resolved }: ReturnType<typeof make>) => { resolved.numbering_source!.relationships_sha256 = numberingDigest },
      ({ resolved }: ReturnType<typeof make>) => { resolved.numbering_source!.relationship_target = 'substituted.xml' },
      ({ resolved }: ReturnType<typeof make>) => { resolved.numbering_source!.relationship_id = 'rIdSubstituted' },
    ]) {
      const value = make()
      mutate(value)
      const shaped = await shapeNativeDocxLinesV1(request(value.document, value.resolved), fakeProviders([]))
      expect(shaped.ok).toBe(false)
    }
  })

  it('emits no partial shaped list when a later source marker is refused', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'first')
    appendParagraph(document, 'paragraph:malformed', 'run:malformed', 'malformed')
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { indent_start_twips: 720, hanging_twips: 720 }
    resolved.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, '1.')
    resolved.paragraphs.push({ paragraph_id: 'paragraph:malformed', applied_styles: [], properties: {}, paragraph_mark_properties: properties() })
    resolved.runs.push({ run_id: 'run:malformed', paragraph_id: 'paragraph:malformed', applied_paragraph_styles: [], applied_character_styles: [], properties: properties() })
    resolved.diagnostics.push({ code: 'MALFORMED_NUMBERING_TEXT', severity: 'unsupported', scope_id: 'paragraph:malformed', preservation: 'preserve-verbatim', message: 'multi-digit placeholder is ambiguous' })
    attestNumbering(document, resolved)
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'unresolved-layout-diagnostic', scope_id: 'paragraph:malformed' })]))
  })

  it('passes UAX #9 Arabic/Latin itemization to the shaper and emits visual cluster order', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'مرحبا A')
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { bidi: true, alignment: 'left', indent_start_twips: 10, indent_end_twips: 20 }
    resolved.runs[0]!.properties = properties({ language: 'ar-SA' })
    const value = request(document, resolved)
    value.available_width_millipoints = 20_000
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(value, fakeProviders(calls))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(calls).toEqual(expect.arrayContaining([
      expect.objectContaining({ text: 'مرحبا', script: 'Arab', direction: 'rtl' }),
      expect.objectContaining({ text: 'A', script: 'Latn', direction: 'ltr' }),
    ]))
    expect(result.value.paragraphs[0]).toMatchObject({ direction: 'rtl', alignment: 'left', indent_start_millipoints: 500, indent_end_millipoints: 1_000 })
    const line = result.value.paragraphs[0]!.lines[0]!
    expect(line.fragments.map((fragment) => fragment.text).join('')).toBe('A ابحرم')
    expect(line.logical_to_visual).toEqual([6, 5, 4, 3, 2, 1, 0])
    expect(result.value.diagnostics).toEqual([])
  })

  it('honors explicit run direction while preserving numeral order and inverse cluster maps', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'abc ')
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs.push({ ...structuredClone(paragraph.runs[0]!), id: 'run:explicit-rtl', text: 'אב 12' })
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { bidi: false, alignment: 'start' }
    resolved.runs[1]!.properties = properties({ rtl: true, language: 'he-IL' })
    const value = request(document, resolved)
    value.available_width_millipoints = 30_000
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(value, fakeProviders(calls))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(calls).toEqual(expect.arrayContaining([
      expect.objectContaining({ text: 'אב', script: 'Hebr', direction: 'rtl' }),
      expect.objectContaining({ text: '12', script: 'Zyyy', direction: 'ltr' }),
    ]))
    const line = result.value.paragraphs[0]!.lines[0]!
    expect(line.fragments.map((fragment) => fragment.text).join('')).toBe('abc 12 בא')
    expect(line.logical_to_visual).toHaveLength(line.fragments.length)
    for (const [logical, visual] of line.logical_to_visual.entries()) expect(line.fragments[visual]!.logical_order).toBe(logical)
    expect(decodeNativeDocxShapedLines(result.value).ok).toBe(true)
    const badMap = structuredClone(result.value)
    badMap.paragraphs[0]!.lines[0]!.logical_to_visual[0] = badMap.paragraphs[0]!.lines[0]!.logical_to_visual[1]!
    expect(decodeNativeDocxShapedLines(badMap).ok).toBe(false)
    const badLevel = structuredClone(result.value)
    badLevel.paragraphs[0]!.lines[0]!.fragments[0]!.bidi_level ^= 1
    expect(decodeNativeDocxShapedLines(badLevel).ok).toBe(false)
    const badProvenance = structuredClone(result.value)
    badProvenance.providers.bidi_revision = `sha256:${'0'.repeat(64)}`
    expect(decodeNativeDocxShapedLines(badProvenance).ok).toBe(false)
  })

  it('expands only bounded U+0020 opportunities and exactly fills soft justified lines', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'aa bb cc')
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { alignment: 'both' }
    const value = request(document, resolved)
    value.available_width_millipoints = 7_000
    const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    const paragraph = result.value.paragraphs[0]!
    expect(paragraph.lines.map((line) => line.fragments.map((fragment) => fragment.text).join(''))).toEqual(['aa bb ', 'cc'])
    expect(paragraph.lines[0]).toMatchObject({ justified: true, available_width_millipoints: 7_000, advance_inline_millipoints: 7_000 })
    expect(paragraph.lines[0]!.fragments.filter((fragment) => fragment.text === ' ').map((fragment) => fragment.justification_expansion_millipoints)).toEqual([1_000, 0])
    const expanded = paragraph.lines[0]!.fragments.find((fragment) => fragment.justification_expansion_millipoints > 0)!
    expect(expanded.glyphs.reduce((sum, glyph) => sum + glyph.advance_x_millipoints, 0)).toBe(expanded.advance_inline_millipoints)
    expect(paragraph.lines[1]).toMatchObject({ justified: false, advance_inline_millipoints: 2_000 })
    expect(decodeNativeDocxShapedLines(result.value).ok).toBe(true)
    const expandedNonSpace = structuredClone(result.value)
    expandedNonSpace.paragraphs[0]!.lines[0]!.fragments.find((fragment) => fragment.justification_expansion_millipoints > 0)!.text = '\u00a0'
    expect(decodeNativeDocxShapedLines(expandedNonSpace).ok).toBe(false)
    const badGlyphSum = structuredClone(result.value)
    badGlyphSum.paragraphs[0]!.lines[0]!.fragments.find((fragment) => fragment.glyphs.length > 0)!.glyphs[0]!.advance_x_millipoints += 1
    expect(decodeNativeDocxShapedLines(badGlyphSum).ok).toBe(false)
    const shiftedExpansion = structuredClone(result.value)
    const shiftedSpaces = shiftedExpansion.paragraphs[0]!.lines[0]!.fragments.filter((fragment) => fragment.text === ' ')
    const amount = shiftedSpaces[0]!.justification_expansion_millipoints
    shiftedSpaces[0]!.justification_expansion_millipoints = 0
    shiftedSpaces[0]!.advance_inline_millipoints -= amount
    shiftedSpaces[0]!.glyphs.at(-1)!.advance_x_millipoints -= amount
    shiftedSpaces[1]!.justification_expansion_millipoints = amount
    shiftedSpaces[1]!.advance_inline_millipoints += amount
    shiftedSpaces[1]!.glyphs.at(-1)!.advance_x_millipoints += amount
    expect(decodeNativeDocxShapedLines(shiftedExpansion).ok).toBe(false)
  })

  it('assigns non-divisible justification remainders in RTL visual order and aligns the final line to logical start', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'א א א א א')
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { bidi: true, alignment: 'both' }
    resolved.runs[0]!.properties = properties({ language: 'he-IL' })
    const value = request(document, resolved)
    value.available_width_millipoints = 7_501
    const first = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    const second = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    expect(first.ok).toBe(true)
    expect(JSON.stringify(second)).toBe(JSON.stringify(first))
    if (!first.ok) return
    const [wrapped, final] = first.value.paragraphs[0]!.lines
    expect(wrapped).toMatchObject({ justified: true, advance_inline_millipoints: 7_501, inline_offset_millipoints: 0 })
    expect(wrapped!.fragments.filter((fragment) => fragment.justification_expansion_millipoints > 0).map((fragment) => fragment.justification_expansion_millipoints)).toEqual([751, 750])
    expect(final).toMatchObject({ justified: false, advance_inline_millipoints: 3_000, inline_offset_millipoints: 4_501 })
  })

  it('does not expand RTL hard-break lines and refuses unqualified justification modes/opportunities', async () => {
    const hardBreakDocument = nativeDocument('א א')
    hardBreakDocument.body.blocks[0]!.paragraph!.runs[2]!.text = 'אב'
    const hardBreakResolved = resolvedLayout(hardBreakDocument)
    hardBreakResolved.paragraphs[0]!.properties = { bidi: true, alignment: 'both' }
    hardBreakResolved.runs.forEach((run) => { run.properties = properties({ language: 'he-IL' }) })
    const hardBreakRequest = request(hardBreakDocument, hardBreakResolved)
    hardBreakRequest.available_width_millipoints = 10_000
    const hardBreak = await shapeNativeDocxLinesV1(hardBreakRequest, fakeProviders([]))
    expect(hardBreak.ok).toBe(true)
    if (hardBreak.ok) expect(hardBreak.value.paragraphs[0]!.lines[0]).toMatchObject({ justified: false, inline_offset_millipoints: 7_000, hard_break_after: { source_run_id: 'run:intro:break' } })

    const noOpportunityDocument = nativeDocument()
    makeTextOnly(noOpportunityDocument, 'aaaa bbbb c')
    const noOpportunityResolved = resolvedLayout(noOpportunityDocument)
    noOpportunityResolved.paragraphs[0]!.properties = { alignment: 'both' }
    const noOpportunityRequest = request(noOpportunityDocument, noOpportunityResolved)
    noOpportunityRequest.available_width_millipoints = 6_500
    const noOpportunity = await shapeNativeDocxLinesV1(noOpportunityRequest, fakeProviders([]))
    expect(noOpportunity.ok && noOpportunity.value.paragraphs).toEqual([])
    expect(noOpportunity.ok && noOpportunity.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'justification-unsupported' })]))

    const distributedDocument = nativeDocument()
    makeTextOnly(distributedDocument, 'abc')
    const distributedResolved = resolvedLayout(distributedDocument)
    distributedResolved.paragraphs[0]!.properties = { alignment: 'distribute' }
    const distributed = await shapeNativeDocxLinesV1(request(distributedDocument, distributedResolved), fakeProviders([]))
    expect(distributed.ok && distributed.value.paragraphs).toEqual([])
    expect(distributed.ok && distributed.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'justification-unsupported' })]))
  })

  it('keeps physical left/right fixed and maps logical start/end through LTR and RTL direction', async () => {
    const expected = { ltr: { left: 500, right: 2_000, start: 500, end: 2_000 }, rtl: { left: 1_000, right: 2_500, start: 2_500, end: 1_000 } } as const
    for (const direction of ['ltr', 'rtl'] as const) for (const alignment of ['left', 'right', 'start', 'end'] as const) {
      const document = nativeDocument()
      makeTextOnly(document, 'ab')
      const resolved = resolvedLayout(document)
      resolved.paragraphs[0]!.properties = { bidi: direction === 'rtl', alignment, indent_start_twips: 10, indent_end_twips: 20 }
      const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs[0]!.lines[0]!.inline_offset_millipoints, `${direction}:${alignment}`).toBe(expected[direction][alignment])
      const forged = structuredClone(result.value)
      forged.paragraphs[0]!.lines[0]!.inline_offset_millipoints += 1
      expect(decodeNativeDocxShapedLines(forged).ok).toBe(false)
    }
  })

  it('combines the tallest ascent, deepest descent, and largest gap across faces', async () => {
    const document = nativeDocument()
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs = [paragraph.runs[0]!, { ...paragraph.runs[2]!, id: 'run:deep', text: 'g' }]
    paragraph.runs[0]!.text = 'A'
    const resolved = resolvedLayout(document)
    resolved.runs[0]!.properties = properties({ font_family: 'Tall' })
    resolved.runs[1]!.properties = properties({ font_family: 'Deep' })
    resolved.paragraphs[0]!.properties = {}
    const makeFace = (id: string, family: string): ResolvedFontFace => ({
      faceId: id,
      family,
      weight: 400,
      style: 'normal',
      stretch: 100,
      sourceKind: 'bundled',
      resourceId: `fonts/${id}.ttf`,
      contentDigest: 'sha256:4bf5122f344554c53bde2ebb8cd2b7e3d1600ad631c385a5d7cce23c7785459a',
      resolution: 'exact',
      matchedFamily: family,
    })
    const tall = makeFace('tall.regular', 'Tall')
    const deep = makeFace('deep.regular', 'Deep')
    const faces = new Map([[tall.family, tall], [deep.family, deep]])
    const value = request(document, resolved)
    value.available_width_millipoints = 20_000
    value.font_manifest = {
      version: 1,
      manifestId: 'mixed.metrics',
      revision: '1',
      faces: [...faces.values()].map((entry) => ({ faceId: entry.faceId, family: entry.family, weight: entry.weight, style: entry.style, stretch: entry.stretch, source: { kind: 'bundled', resourceId: entry.resourceId, contentDigest: entry.contentDigest } })),
      fallbackChains: [],
    }
    const providers = {
      resolver: {
        providerId: 'mixed-resolver', providerRevision: '1',
        resolve: ({ run }: Parameters<NativeFontResolver['resolve']>[0]) => ({ status: 'resolved' as const, face: faces.get(run.font.families[0]!)!, attemptedFaceIds: [], decisions: [] }),
        load: (loadedFace: ResolvedFontFace): FontResource => ({ face: loadedFace, bytes: new Uint8Array([1]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 0 } }),
      },
      shaper: {
        providerId: 'mixed-shaper', providerRevision: '1',
        shape: ({ run, font }: Parameters<NativeTextShaper['shape']>[0]): ShapedSegment => {
          const line = font.face.faceId === tall.faceId
            ? { ascentMilliPoints: 9_000, descentMilliPoints: -1_000, lineGapMilliPoints: 500, lineHeightMilliPoints: 10_500 }
            : { ascentMilliPoints: 7_000, descentMilliPoints: -4_000, lineGapMilliPoints: 2_000, lineHeightMilliPoints: 13_000 }
          return { startUtf16: 0, endUtf16: 1, face: font.face, glyphs: [{ glyphId: 1, clusterIndex: 0, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 }], clusters: [{ startUtf16: 0, endUtf16: 1, glyphStart: 0, glyphEnd: 1, advanceInlineMilliPoints: 1_000 }], metrics: { fontSizeMilliPoints: run.fontSizeMilliPoints, ...line }, advanceInlineMilliPoints: 1_000, advanceBlockMilliPoints: 0 }
        },
      },
    } satisfies { resolver: NativeFontResolver; shaper: NativeTextShaper }
    const result = await shapeNativeDocxLinesV1(value, providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs[0]!.lines[0]).toMatchObject({ ascent_millipoints: 9_000, descent_millipoints: -4_000, line_gap_millipoints: 2_000, line_height_millipoints: 15_000 })
  })

  it('atomically refuses astral emoji before itemization or provider calls', async () => {
    const document = nativeDocument()
    makeTextOnly(document, '漢字😀文')
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = {}
    const value = request(document, resolved)
    value.available_width_millipoints = 2_000
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(value, fakeProviders(calls))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'bidi-resolution-refusal', message: expect.stringContaining('unsupported-scalar') })]))
    expect(calls).toEqual([])

    makeTextOnly(document, 'א\u{1E900}ב')
    const adlamCalls: typeof calls = []
    const adlam = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), fakeProviders(adlamCalls))
    expect(adlam).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ paragraphs: [], diagnostics: expect.arrayContaining([expect.objectContaining({ code: 'bidi-resolution-refusal' })]) }) }))
    expect(adlamCalls).toEqual([])
  })

  it('keeps non-breaking spaces glued and emits an overfull cluster group instead of splitting it', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'a\u00a0b')
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = {}
    const value = request(document, resolved)
    value.available_width_millipoints = 2_000
    const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs[0]!.lines).toHaveLength(1)
    expect(result.value.paragraphs[0]!.lines[0]).toMatchObject({ advance_inline_millipoints: 3_000 })
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'cluster-overflow' })]))
  })

  it('retains trailing and consecutive hard breaks as explicit empty lines', async () => {
    const document = nativeDocument()
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs = [
      { ...paragraph.runs[1]!, kind: 'control', id: 'run:break:1', control: 'line-break', reference: undefined },
      { ...paragraph.runs[2]!, kind: 'control', id: 'run:break:2', control: 'line-break', text: undefined, drawing: undefined },
    ]
    const resolved = resolvedLayout(document)
    resolved.runs = paragraph.runs.map((run) => ({ run_id: run.id, paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: properties() }))
    resolved.paragraphs[0]!.properties = { line: 240, line_rule: 'exact' }
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
    expect(result.ok, result.ok ? '' : JSON.stringify(result.issues)).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs[0]!.lines).toHaveLength(3)
    expect(result.value.paragraphs[0]!.lines.map((line) => line.hard_break_after?.source_run_id)).toEqual(['run:break:1', 'run:break:2', undefined])
    expect(result.value.paragraphs[0]!.lines.every((line) => line.line_height_millipoints === 12_000)).toBe(true)
    expect(result.value.paragraphs[0]!.lines.every((line) => line.fragments.length === 0)).toBe(true)
  })

  it('uses paragraph-mark font metrics for empty and hidden-only paragraphs without fake fragments', async () => {
    for (const kind of ['empty', 'hidden'] as const) {
      const document = nativeDocument()
      const paragraph = document.body.blocks[0]!.paragraph!
      if (kind === 'empty') paragraph.runs = []
      else {
        paragraph.runs = [{ ...paragraph.runs[0]!, text: 'hidden source' }]
      }
      const resolved = resolvedLayout(document)
      if (kind === 'hidden') resolved.runs[0]!.properties = properties({ hidden: true })
      const requests: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
      const providers = fakeProviders(requests)
      let shapeCalls = 0
      const shape = providers.shaper.shape.bind(providers.shaper)
      providers.shaper.shape = (providerRequest) => { shapeCalls += 1; return shape(providerRequest) }
      const result = await shapeNativeDocxLinesV1(request(document, resolved), providers)
      expect(result.ok, kind).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs, kind).toHaveLength(1)
      expect(result.value.paragraphs[0]!.lines, kind).toHaveLength(1)
      expect(result.value.paragraphs[0]!.lines[0], kind).toMatchObject({ line_height_millipoints: 12_000, fragments: [] })
      expect(requests, kind).toEqual([expect.objectContaining({ text: '', family: 'Carlito', size: 10_000, script: 'Zyyy' })])
      expect(shapeCalls, kind).toBe(0)
    }
  })

  it('atomically refuses a drawing paragraph and skips tables with scoped diagnostics', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'outside')
    const drawingRun = structuredClone(fixture.body.blocks[0]!.paragraph!.runs[2]!)
    document.body.blocks[0]!.paragraph!.runs.push(drawingRun)
    document.body.blocks.push(structuredClone(fixture.body.blocks[1]!))
    document.passthrough_parts = fixture.passthrough_parts.filter((part) => part.part_name === drawingRun.drawing?.media_part)
    const resolved = resolvedLayout(document)
    const table = document.body.blocks[1]!.table!
    const cellParagraph = table.rows[0]!.cells[0]!.paragraphs[0]!
    resolved.tables.push({ table_id: table.id, style_id: table.table_style_id })
    resolved.paragraphs.push({ paragraph_id: cellParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: properties() })
    resolved.runs.push({ run_id: cellParagraph.runs[0]!.id, paragraph_id: cellParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
    expect(result.ok, result.ok ? '' : JSON.stringify(result.issues)).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'drawing-layout-unsupported', source_id: 'run:intro:drawing' }),
      expect.objectContaining({ code: 'table-layout-unsupported', scope_id: 'table:summary' }),
    ]))
  })

  it('does not let mere numbering provenance turn unrelated drawing/table refusals into a document-wide erase', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'drawing owner')
    document.body.blocks[0]!.paragraph!.runs.push(structuredClone(fixture.body.blocks[0]!.paragraph!.runs[2]!))
    appendParagraph(document, 'paragraph:safe', 'run:safe', 'safe')
    document.body.blocks.push(structuredClone(fixture.body.blocks[1]!))
    const resolved = resolvedLayout(document)
    resolved.paragraphs.push({ paragraph_id: 'paragraph:safe', applied_styles: [], properties: {}, paragraph_mark_properties: properties() })
    resolved.runs.push({ run_id: 'run:safe', paragraph_id: 'paragraph:safe', applied_paragraph_styles: [], applied_character_styles: [], properties: properties() })
    const table = document.body.blocks[2]!.table!
    const cellParagraph = table.rows[0]!.cells[0]!.paragraphs[0]!
    resolved.tables.push({ table_id: table.id })
    resolved.paragraphs.push({ paragraph_id: cellParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: properties() })
    resolved.runs.push({ run_id: cellParagraph.runs[0]!.id, paragraph_id: cellParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: properties() })
    attestNumbering(document, resolved)
    document.passthrough_parts.push(structuredClone(fixture.passthrough_parts.find((part) => part.part_name === document.body.blocks[0]!.paragraph!.runs[1]!.drawing!.media_part)!))
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
    expect(result.ok, result.ok ? '' : JSON.stringify(result.issues)).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs.map((paragraph) => paragraph.paragraph_id)).toEqual(['paragraph:safe'])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'drawing-layout-unsupported', scope_id: 'paragraph:intro' }),
      expect.objectContaining({ code: 'table-layout-unsupported', scope_id: table.id }),
    ]))
  })

  it('blocks theme/script-dependent runs and missing metrics instead of guessing', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'unresolved')
    const resolved = resolvedLayout(document)
    resolved.diagnostics.push({ code: 'THEME_FONT_PRESERVED', severity: 'unsupported', scope_id: 'run:intro:text', preservation: 'preserve-verbatim', message: 'theme font unresolved' })
    delete resolved.runs[0]!.properties.font_size_half_points
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders(calls))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(calls).toEqual([])
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'unresolved-layout-diagnostic', source_id: 'run:intro:text', source_diagnostic_code: 'THEME_FONT_PRESERVED' })]))
  })

  it('fails closed on unresolved layout diagnostics but allows documented paint-only color gaps', async () => {
    for (const code of ['INVALID_PARAGRAPH_INDENT', 'INCOMPLETE_LINE_SPACING', 'PICTURE_BULLET_PRESERVED']) {
      const document = nativeDocument()
      makeTextOnly(document, 'blocked')
      const resolved = resolvedLayout(document)
      resolved.diagnostics.push({ code, severity: 'unsupported', scope_id: 'paragraph:intro', preservation: 'preserve-verbatim', message: `${code} fixture` })
      const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
      const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders(calls))
      expect(result.ok).toBe(true)
      if (!result.ok) continue
      expect(calls, code).toEqual([])
      expect(result.value.paragraphs, code).toEqual([])
      expect(result.value.diagnostics, code).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'unresolved-layout-diagnostic', source_diagnostic_code: code, source_diagnostic_message: `${code} fixture` })]))
    }

    const document = nativeDocument()
    makeTextOnly(document, 'painted')
    const resolved = resolvedLayout(document)
    resolved.diagnostics.push({ code: 'THEME_COLOR_PRESERVED', severity: 'unsupported', scope_id: 'run:intro:text', preservation: 'preserve-verbatim', message: 'theme color retained' })
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders(calls))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(calls.length).toBeGreaterThan(0)
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'paint-diagnostic-preserved', source_diagnostic_code: 'THEME_COLOR_PRESERVED' })]))
  })

  it('allows exact latent style behavior metadata only at its bound document scope', async () => {
    for (const variant of ['qualified', 'old-code', 'run-scope', 'wrong-part', 'wrong-path', 'active-style'] as const) {
      const document = nativeDocument()
      makeTextOnly(document, 'Latent metadata does not format text')
      const resolved = resolvedLayout(document)
      resolved.source_parts.styles_part = 'word/styles.xml'
      resolved.diagnostics.push({
        code: variant === 'old-code' ? 'LATENT_STYLES_PRESERVED' : variant === 'active-style' ? 'MISSING_PARAGRAPH_STYLE' : 'LATENT_STYLE_BEHAVIOR_PRESERVED',
        severity: 'unsupported', preservation: 'preserve-verbatim',
        scope_id: variant === 'run-scope' ? 'run:intro:text' : document.document_id,
        part_name: variant === 'wrong-part' ? 'word/document.xml' : 'word/styles.xml',
        path: variant === 'wrong-path' ? '/w:styles[1]/w:style[1]' : '/w:styles[1]/w:latentStyles[1]',
        message: 'UI metadata is preserved',
      })
      const before = JSON.stringify(resolved)
      const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
      expect(result.ok, variant).toBe(true)
      if (!result.ok) continue
      if (variant === 'qualified') {
        expect(result.value.paragraphs.length).toBeGreaterThan(0)
        expect(result.value.diagnostics).toEqual([])
      } else expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'unresolved-layout-diagnostic' })]))
      expect(JSON.stringify(resolved)).toBe(before)
    }
  })

  it('preserves qualified font descriptors and requires exact supplied faces', async () => {
    for (const variant of ['qualified', 'unknown', 'wrong-part', 'wrong-path', 'substitute', 'fallback'] as const) {
      const document = nativeDocument()
      makeTextOnly(document, 'Font descriptors do not supply glyphs')
      const resolved = resolvedLayout(document)
      resolved.source_parts.font_table_part = 'word/fonts.xml'
      resolved.diagnostics.push({
        code: variant === 'unknown' ? 'UNMODELED_FONT_METADATA' : 'FONT_MATCHING_METADATA_PRESERVED',
        severity: 'unsupported', preservation: 'preserve-verbatim', scope_id: document.document_id,
        part_name: variant === 'wrong-part' ? 'word/document.xml' : 'word/fonts.xml',
        path: variant === 'wrong-path' ? '/w:fonts[1]/w:font[1]/w:unknown[1]' : '/w:fonts[1]/w:font[1]/w:panose1[1]',
        message: 'Validated matching hints retained',
      })
      const providers = fakeProviders([])
      if (variant === 'substitute' || variant === 'fallback') providers.resolver.resolve = () => ({ status: 'resolved', face: { ...face, resolution: variant }, attemptedFaceIds: [face.faceId], decisions: [] })
      const before = JSON.stringify(resolved)
      const result = await shapeNativeDocxLinesV1(request(document, resolved), providers)
      expect(result.ok, variant).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs.length > 0, variant).toBe(variant === 'qualified')
      if (variant === 'substitute' || variant === 'fallback') expect(result.value.diagnostics.some((entry) => entry.code === 'provider-refusal')).toBe(true)
      expect(JSON.stringify(resolved)).toBe(before)
    }
  })

  it('emits pagination-policy diagnostics without attempting pagination', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'policy')
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { keep_next: true, keep_lines: false, page_break_before: true, widow_control: false }
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    const messages = result.value.diagnostics.filter((diagnostic) => diagnostic.code === 'page-control-deferred').map((diagnostic) => diagnostic.message)
    expect(messages).toHaveLength(4)
    expect(messages.join(' ')).toContain('page_break_before')
  })

  it('atomically omits a multi-run paragraph when a later provider refuses', async () => {
    const document = nativeDocument()
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs = [
      { ...paragraph.runs[0]!, text: 'surviving prefix' },
      { ...paragraph.runs[2]!, id: 'run:refused-tail', text: 'refused tail' },
    ]
    const resolved = resolvedLayout(document)
    const providers = fakeProviders([])
    const resolve = providers.resolver.resolve.bind(providers.resolver)
    providers.resolver.resolve = (providerRequest) => providerRequest.run.text.includes('refused')
      ? { status: 'refused', attemptedFaceIds: [face.faceId], decisions: [{ code: 'missing-glyph', message: 'tail cannot be shaped', recoverable: false }] }
      : resolve(providerRequest)
    const result = await shapeNativeDocxLinesV1(request(document, resolved), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'provider-refusal', source_id: 'run:refused-tail' })]))
  })

  it('stops span and run provider calls immediately after the first blocking failure', async () => {
    const document = nativeDocument()
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs = [
      { ...paragraph.runs[0]!, text: 'AاAا' },
      { ...paragraph.runs[2]!, id: 'run:must-not-resolve', text: 'later run' },
    ]
    const resolved = resolvedLayout(document)
    let resolveCalls = 0
    const providers = fakeProviders([])
    providers.resolver.resolve = () => {
      resolveCalls += 1
      return { status: 'refused', attemptedFaceIds: [face.faceId], decisions: [{ code: 'font-not-found', message: 'first span refused', recoverable: false }] }
    }
    const result = await shapeNativeDocxLinesV1(request(document, resolved), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(resolveCalls).toBe(1)
    expect(result.value.paragraphs).toEqual([])
  })

  it('only allows table diagnostics on table scopes and propagates skipped-table descendant diagnostics', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'outside')
    document.body.blocks.push(structuredClone(fixture.body.blocks[1]!))
    const resolved = resolvedLayout(document)
    const table = document.body.blocks[1]!.table!
    const cellParagraph = table.rows[0]!.cells[0]!.paragraphs[0]!
    const cellRun = cellParagraph.runs[0]!
    resolved.tables.push({ table_id: table.id })
    resolved.paragraphs.push({ paragraph_id: cellParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: properties() })
    resolved.runs.push({ run_id: cellRun.id, paragraph_id: cellParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: properties() })
    resolved.diagnostics.push(
      { code: 'MISSING_TABLE_STYLE', severity: 'unsupported', scope_id: 'paragraph:intro', preservation: 'preserve-verbatim', message: 'maliciously mis-scoped table code' },
      { code: 'TABLE_STYLE_EFFECTS_PRESERVED', severity: 'unsupported', scope_id: cellRun.id, preservation: 'preserve-verbatim', message: 'cell run table effects retained' },
    )
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders(calls))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(calls).toEqual([])
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'unresolved-layout-diagnostic', scope_id: 'paragraph:intro', source_diagnostic_code: 'MISSING_TABLE_STYLE' }),
      expect.objectContaining({ code: 'table-layout-unsupported', scope_id: table.id, source_diagnostic_code: 'TABLE_STYLE_EFFECTS_PRESERVED', source_diagnostic_message: 'cell run table effects retained' }),
    ]))
  })

  it('allows paint diagnostics only on paragraph/run scopes', async () => {
    const documentScoped = nativeDocument()
    makeTextOnly(documentScoped, 'blocked')
    const documentResolved = resolvedLayout(documentScoped)
    documentResolved.diagnostics.push({ code: 'INVALID_COLOR', severity: 'unsupported', scope_id: documentScoped.document_id, preservation: 'preserve-verbatim', message: 'mis-scoped document paint gap' })
    const documentCalls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const blocked = await shapeNativeDocxLinesV1(request(documentScoped, documentResolved), fakeProviders(documentCalls))
    expect(blocked.ok).toBe(true)
    if (!blocked.ok) return
    expect(documentCalls).toEqual([])
    expect(blocked.value.paragraphs).toEqual([])
    expect(blocked.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'unresolved-layout-diagnostic', source_diagnostic_code: 'INVALID_COLOR' })]))
    expect(blocked.value.diagnostics).not.toEqual(expect.arrayContaining([expect.objectContaining({ code: 'paint-diagnostic-preserved', source_diagnostic_code: 'INVALID_COLOR' })]))

    const tableScoped = nativeDocument()
    makeTextOnly(tableScoped, 'outside')
    tableScoped.body.blocks.push(structuredClone(fixture.body.blocks[1]!))
    const tableResolved = resolvedLayout(tableScoped)
    const table = tableScoped.body.blocks[1]!.table!
    const cellParagraph = table.rows[0]!.cells[0]!.paragraphs[0]!
    tableResolved.tables.push({ table_id: table.id })
    tableResolved.paragraphs.push({ paragraph_id: cellParagraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: properties() })
    tableResolved.runs.push({ run_id: cellParagraph.runs[0]!.id, paragraph_id: cellParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: properties() })
    tableResolved.diagnostics.push({ code: 'INVALID_COLOR', severity: 'unsupported', scope_id: table.id, preservation: 'preserve-verbatim', message: 'mis-scoped table paint gap' })
    const tableResult = await shapeNativeDocxLinesV1(request(tableScoped, tableResolved), fakeProviders([]))
    expect(tableResult.ok).toBe(true)
    if (!tableResult.ok) return
    expect(tableResult.value.paragraphs.map((paragraph) => paragraph.paragraph_id)).toEqual(['paragraph:intro'])
    expect(tableResult.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'table-layout-unsupported', scope_id: table.id, source_diagnostic_code: 'INVALID_COLOR' })]))
    expect(tableResult.value.diagnostics).not.toEqual(expect.arrayContaining([expect.objectContaining({ code: 'paint-diagnostic-preserved', source_diagnostic_code: 'INVALID_COLOR' })]))
  })

  it('defines tab stops relative to the indented paragraph-content origin', async () => {
    const document = nativeDocument()
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs = [
      { ...paragraph.runs[0]!, text: 'a' },
      { ...paragraph.runs[1]!, kind: 'control', id: 'run:tab', control: 'tab', reference: undefined },
      { ...paragraph.runs[2]!, id: 'run:after-tab', text: 'b' },
    ]
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { indent_left_twips: 10 }
    const value = request(document, resolved)
    value.available_width_millipoints = 10_000
    value.tab_interval_millipoints = 4_000
    const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    const line = result.value.paragraphs[0]!.lines[0]!
    expect(line.inline_offset_millipoints).toBe(500)
    expect(line.fragments.find((fragment) => fragment.source_id === 'run:tab')).toMatchObject({ source_kind: 'tab', advance_inline_millipoints: 3_000 })
  })

  it('refuses the complete shaped paragraph projection when a skipped table contains numbered paragraphs', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'first')
    document.body.blocks.push(structuredClone(fixture.body.blocks[1]!))
    appendParagraph(document, 'paragraph:after-table', 'run:after-table', 'after')
    document.passthrough_parts = []
    const resolved = resolvedLayout(document)
    const table = document.body.blocks[1]!.table!
    const cellParagraph = table.rows[0]!.cells[0]!.paragraphs[0]!
    resolved.tables.push({ table_id: table.id })
    resolved.paragraphs.push({ paragraph_id: cellParagraph.id, applied_styles: [], properties: { indent_start_twips: 720, hanging_twips: 720 }, paragraph_mark_properties: properties(), numbering: resolvedNumbering(cellParagraph.id, 2, '2.', 'space') })
    resolved.runs.push({ run_id: cellParagraph.runs[0]!.id, paragraph_id: cellParagraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: properties() })
    resolved.paragraphs.push({ paragraph_id: 'paragraph:after-table', applied_styles: [], properties: { indent_start_twips: 720, hanging_twips: 720 }, paragraph_mark_properties: properties(), numbering: resolvedNumbering('paragraph:after-table', 3, '3.', 'space') })
    resolved.runs.push({ run_id: 'run:after-table', paragraph_id: 'paragraph:after-table', applied_paragraph_styles: [], applied_character_styles: [], properties: properties() })
    resolved.paragraphs[0]!.properties = { indent_start_twips: 720, hanging_twips: 720 }
    resolved.paragraphs[0]!.numbering = resolvedNumbering('paragraph:intro', 1, '1.', 'space')
    attestNumbering(document, resolved)
    const value = request(document, resolved)
    value.available_width_millipoints = 30_000
    const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'table-layout-unsupported', scope_id: table.id })]))
  })

  it('rejects inconsistent injected clusters and returns deterministic output for valid providers', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'safe')
    const value = request(document, resolvedLayout(document))
    const good = fakeProviders([])
    const first = await shapeNativeDocxLinesV1(value, good)
    const second = await shapeNativeDocxLinesV1(structuredClone(value), fakeProviders([]))
    expect(first).toEqual(second)

    const bad = fakeProviders([])
    bad.shaper.shape = ({ run }) => ({
      startUtf16: 0,
      endUtf16: run.text.length,
      face,
      glyphs: [],
      clusters: [{ startUtf16: 1, endUtf16: run.text.length, glyphStart: 0, glyphEnd: 0, advanceInlineMilliPoints: 0 }],
      metrics: { fontSizeMilliPoints: 10_000, ascentMilliPoints: 8_000, descentMilliPoints: -2_000, lineGapMilliPoints: 2_000, lineHeightMilliPoints: 12_000 },
      advanceInlineMilliPoints: 0,
      advanceBlockMilliPoints: 0,
    })
    const refused = await shapeNativeDocxLinesV1(value, bad)
    expect(refused.ok).toBe(true)
    if (!refused.ok) return
    expect(refused.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output', source_id: 'run:intro:text' })]))
  })

  it('rejects malicious resolver, resource, face, metrics, and coordinate outputs', async () => {
    const cases: Array<[string, (providers: { resolver: NativeFontResolver; shaper: NativeTextShaper }) => void]> = [
      ['unmanifested resolver face', (providers) => {
        providers.resolver.resolve = () => ({ status: 'resolved', face: { ...face, faceId: 'missing.face' }, attemptedFaceIds: [], decisions: [] })
      }],
      ['unbounded resolver decision metadata', (providers) => {
        providers.resolver.resolve = () => ({ status: 'resolved', face, attemptedFaceIds: [], decisions: [{ code: 'missing-glyph', message: 'x'.repeat(4_097), recoverable: false }] })
      }],
      ['invalid attempted face metadata', (providers) => {
        providers.resolver.resolve = () => ({ status: 'resolved', face, attemptedFaceIds: ['bad face id'], decisions: [] })
      }],
      ['mismatched loaded resource', (providers) => {
        providers.resolver.load = () => ({ face: { ...face, resourceId: 'fonts/other.ttf' }, bytes: new Uint8Array([1]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 0 } })
      }],
      ['resource digest mismatch', (providers) => {
        providers.resolver.load = () => ({ face, bytes: new Uint8Array([9]), metrics: { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 0 } })
      }],
      ['invalid resource metric signs', (providers) => {
        providers.resolver.load = () => ({ face, bytes: new Uint8Array([0, 1, 2, 3]), metrics: { unitsPerEm: 1_000, ascender: -1, descender: 1, lineGap: -1 } })
      }],
      ['mismatched shaped face', (providers) => {
        const shape = providers.shaper.shape.bind(providers.shaper)
        providers.shaper.shape = async (providerRequest) => ({ ...(await shape(providerRequest) as ShapedSegment), face: { ...face, matchedFamily: 'Other' } })
      }],
      ['invalid shaped metrics', (providers) => {
        const shape = providers.shaper.shape.bind(providers.shaper)
        providers.shaper.shape = async (providerRequest) => {
          const shaped = await shape(providerRequest) as ShapedSegment
          return { ...shaped, metrics: { ...shaped.metrics, fontSizeMilliPoints: providerRequest.run.fontSizeMilliPoints + 1, ascentMilliPoints: -1, lineHeightMilliPoints: 2_201 } }
        }
      }],
      ['unbounded glyph coordinate', (providers) => {
        const shape = providers.shaper.shape.bind(providers.shaper)
        providers.shaper.shape = async (providerRequest) => {
          const shaped = await shape(providerRequest) as ShapedSegment
          return { ...shaped, glyphs: shaped.glyphs.map((glyph, index) => index === 0 ? { ...glyph, offsetXMilliPoints: 1_000_000_001 } : glyph) }
        }
      }],
    ]
    for (const [name, mutate] of cases) {
      const document = nativeDocument()
      makeTextOnly(document, 'safe')
      const providers = fakeProviders([])
      mutate(providers)
      const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
      expect(result.ok, name).toBe(true)
      if (!result.ok) continue
      expect(result.value.paragraphs, name).toEqual([])
      expect(result.value.diagnostics, name).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'invalid-provider-output' })]))
    }
  })

  it('bounds thrown provider errors before exposing diagnostics', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'safe')
    const providers = fakeProviders([])
    providers.resolver.resolve = () => { throw new Error('x'.repeat(100_000)) }
    const result = await shapeNativeDocxLinesV1(request(document, resolvedLayout(document)), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    const diagnostic = result.value.diagnostics.find((entry) => entry.code === 'provider-failure')
    expect(diagnostic?.message.length).toBeLessThan(4_200)
  })

  it('refuses out-of-bound resolved line heights and paragraph block sums without partial output', async () => {
    const autoDocument = nativeDocument()
    makeTextOnly(autoDocument, 'auto')
    const invalidAuto = resolvedLayout(autoDocument)
    invalidAuto.paragraphs[0]!.properties = { line: 2_401, line_rule: 'auto' }
    const rejected = await shapeNativeDocxLinesV1(request(autoDocument, invalidAuto), fakeProviders([]))
    expect(rejected.ok).toBe(false)
    if (!rejected.ok) expect(rejected.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '/resolved_layout/paragraphs/0/properties/line', code: 'LIMIT_EXCEEDED' })]))

    const document = nativeDocument()
    makeTextOnly(document, 'x'.repeat(101))
    const resolved = resolvedLayout(document)
    resolved.paragraphs[0]!.properties = { line: 200_000_000, line_rule: 'exact' }
    const value = request(document, resolved)
    value.available_width_millipoints = 1
    const result = await shapeNativeDocxLinesV1(value, fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'resource-limit', scope_id: 'paragraph:intro', message: expect.stringContaining('block advance') })]))
  })

  it('enforces the source/cluster budget before another provider allocation and emits no partial paragraph', async () => {
    const document = nativeDocument()
    const paragraph = document.body.blocks[0]!.paragraph!
    paragraph.runs = [
      { ...paragraph.runs[0]!, text: 'a'.repeat(250_001) },
      { ...paragraph.runs[2]!, id: 'run:budget:2', text: 'b'.repeat(250_001) },
    ]
    const resolved = resolvedLayout(document)
    const calls: Array<{ text: string; family: string; size: number; script: string; direction: string }> = []
    const providers = fakeProviders(calls)
    providers.shaper.shape = ({ run }): ShapedSegment => ({
      startUtf16: 0,
      endUtf16: run.text.length,
      face,
      glyphs: [{ glyphId: 1, clusterIndex: 0, advanceXMilliPoints: 1_000, advanceYMilliPoints: 0, offsetXMilliPoints: 0, offsetYMilliPoints: 0 }],
      clusters: [{ startUtf16: 0, endUtf16: run.text.length, glyphStart: 0, glyphEnd: 1, advanceInlineMilliPoints: 1_000 }],
      metrics: { fontSizeMilliPoints: run.fontSizeMilliPoints, ascentMilliPoints: 8_000, descentMilliPoints: -2_000, lineGapMilliPoints: 2_000, lineHeightMilliPoints: 12_000 },
      advanceInlineMilliPoints: 1_000,
      advanceBlockMilliPoints: 0,
    })
    const result = await shapeNativeDocxLinesV1(request(document, resolved), providers)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(calls).toHaveLength(0)
    expect(result.value.paragraphs).toEqual([])
    expect(result.value.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'bidi-resolution-refusal', scope_id: 'paragraph:intro', message: expect.stringContaining('resource-limit') })]))
  })

  it('reserves a final explicit diagnostic-overflow marker', async () => {
    const document = nativeDocument()
    makeTextOnly(document, 'colors')
    const resolved = resolvedLayout(document)
    resolved.diagnostics = Array.from({ length: 1_000 }, (_unused, index) => ({ code: 'THEME_COLOR_PRESERVED', severity: 'unsupported' as const, scope_id: 'run:intro:text', preservation: 'preserve-verbatim' as const, message: `paint diagnostic ${index}` }))
    const result = await shapeNativeDocxLinesV1(request(document, resolved), fakeProviders([]))
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.value.diagnostics).toHaveLength(1_000)
    expect(result.value.diagnostics.at(-1)).toMatchObject({ code: 'diagnostic-overflow', scope_id: document.document_id })
  })

  it('has no browser, DOM, canvas, HTML renderer, or Mammoth imports', () => {
    const source = readFileSync(new URL('./nativeShapingLines.ts', import.meta.url), 'utf8')
    const imports = [...source.matchAll(/from\s+['"]([^'"]+)['"]/g)].map((match) => match[1])
    expect(imports).not.toEqual(expect.arrayContaining(['mammoth', 'canvas', 'react', 'konva']))
    expect(source).not.toMatch(/\b(?:window|HTMLElement|CanvasRenderingContext2D)\b/)
  })
})
