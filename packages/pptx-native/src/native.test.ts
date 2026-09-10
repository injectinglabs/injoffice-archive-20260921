import { readFileSync, readdirSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  PPTX_NATIVE_SCHEMA_SHA256,
  PPTX_NATIVE_RESOURCE_LIMITS,
  assertNativePptx,
  canonicalizeNativePptx,
  stringifyNativePptx,
  validateNativePptx,
} from './index'
import type { NativeElement, NativePptxDeck } from './types'
import { createHash } from 'node:crypto'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const fixtureRoot = resolve(repositoryRoot, 'go/pptxpatch/testdata/native-contract')

function fixture(path: string): unknown {
  return JSON.parse(readFileSync(resolve(fixtureRoot, path), 'utf8'))
}

describe('native PPTX contract', () => {
  it('retains bounded authored markers and refuses contradictory or malformed bullet metadata', () => {
    const deck = fixture('valid/parsed-full.json') as NativePptxDeck
    const element = deck.slides[0]!.elements.find((item) => item.kind === 'text')!
    if (element.kind !== 'text') throw new Error('text fixture missing')
    const paragraph = element.paragraphs[0]!
    Object.assign(paragraph, { bullet: true, bulletCharacter: '▪', marginLeftEmu: 300000, indentEmu: -100000 })
    expect(validateNativePptx(deck).ok).toBe(true)
    for (const marker of ['🙂', '•']) { paragraph.bulletCharacter = marker; expect(validateNativePptx(deck).ok).toBe(true) }
    for (const invalid of [ { bullet: false, bulletCharacter: '▪' }, { bullet: true, bulletCharacter: 'ab' }, { bullet: true, bulletCharacter: '\u0085' }, { bullet: true, bulletCharacter: '▪', marginLeftEmu: -1 }, { bullet: true, bulletCharacter: '▪', marginLeftEmu: 0, indentEmu: 51206401 } ]) {
      Object.assign(paragraph, invalid)
      expect(validateNativePptx(deck).ok).toBe(false)
    }
  })
  it('validates exact picture crop edges and refuses invalid or empty source rectangles', () => {
    for (const crop of [
      { left: 0, top: 0, right: 0, bottom: 0 },
      { left: 12500, top: 25000, right: 37500, bottom: 0 },
      { left: 99999, top: 0, right: 0, bottom: 99999 },
    ]) {
      const deck = fixture('valid/parsed-full.json') as NativePptxDeck
      const picture = deck.slides[0]!.elements.find((item) => item.kind === 'picture')!
      if (picture.kind !== 'picture') throw new Error('fixture requires a picture')
      picture.crop = crop
      expect(validateNativePptx(deck).ok).toBe(true)
      for (const invalid of [
        { ...crop, left: -1 }, { ...crop, top: 0.5 }, { ...crop, right: 100000 },
        { ...crop, left: 50000, right: 50000 }, { ...crop, top: 99999, bottom: 1 },
        { left: 0, top: 0, right: 0 }, { ...crop, unknown: 1 },
      ]) {
        Object.assign(picture, { crop: invalid })
        expect(validateNativePptx(deck).ok, JSON.stringify(invalid)).toBe(false)
      }
    }
  })

  it('accepts the shared authored and parsed fixtures', () => {
    for (const name of readdirSync(resolve(fixtureRoot, 'valid')).filter((item) => item.endsWith('.json'))) {
      expect(validateNativePptx(fixture(`valid/${name}`)), name).toEqual({ ok: true, value: fixture(`valid/${name}`) })
    }
  })

  it('rejects every shared invalid fixture', () => {
    for (const name of readdirSync(resolve(fixtureRoot, 'invalid')).filter((item) => item.endsWith('.json'))) {
      const result = validateNativePptx(fixture(`invalid/${name}`))
      expect(result.ok, name).toBe(false)
      if (!result.ok) expect(result.issues.length, name).toBeGreaterThan(0)
    }
  })

  it('rejects every omitted required field whose valid value may be zero or empty', () => {
    const result = validateNativePptx(fixture('invalid/missing-required-zero-values.json'))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.path)).toEqual(expect.arrayContaining([
      '$.assets[0].byteLength',
      '$.slides[0].elements[0].transform.x',
      '$.slides[0].elements[0].paragraphs[0].runs[0].text',
      '$.slides[0].elements[1].stroke.widthEmu',
      '$.slides[0].elements[2].table.rows[0][0].text',
      '$.slides[0].elements[2].table.rows[0][0].border.widthEmu',
    ]))
  })

  it('rejects authored source claims and dangling picture assets', () => {
    const deck = fixture('valid/authored-minimal.json') as NativePptxDeck
    deck.slides.push({
      id: 'slide-1',
      provenance: 'authored',
      elements: [{
        kind: 'picture', id: 'picture-1', provenance: 'authored', assetId: 'missing',
        transform: { x: 0, y: 0, cx: 1, cy: 1 }, passthrough: [],
        source: { partName: 'ppt/slides/slide1.xml', objectId: 'cNvPr-2', fingerprintSha256: '1'.repeat(64) },
        compatibility: { status: 'editable', diagnostics: [] },
      }],
      passthrough: [], compatibility: { status: 'editable', diagnostics: [] },
    })
    const result = validateNativePptx(deck)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toEqual(expect.arrayContaining(['native.authoredSource', 'native.assetReference']))
  })

  it('rejects metadata-only parsed source assets without a bound read capability', () => {
    const result = validateNativePptx(fixture('invalid/parsed-source-asset-missing-capability.json'))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.assetReadCapability')
  })

  it('allows a missing shape preset only for an explicit refused placeholder', () => {
    const result = validateNativePptx(fixture('invalid/editable-shape-missing-preset.json'))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.shapePreset')
  })

  it('rejects partial native stroke semantics', () => {
    const result = validateNativePptx(fixture('invalid/partial-stroke-semantics.json'))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.strokeMetadata')
  })

  it('enforces the shared DrawingML stroke and miter bounds', () => {
    const result = validateNativePptx(fixture('invalid/oversized-stroke-width.json'))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.path)).toEqual(expect.arrayContaining([
      '$.slides[0].elements[0].stroke.widthEmu',
      '$.slides[0].elements[0].stroke.miterLimit',
      '$.slides[0].elements[1].table.rows[0][0].border.widthEmu',
    ]))
    expect(PPTX_NATIVE_RESOURCE_LIMITS).toMatchObject({ maxLineWidthEmu: 20_116_800, maxDrawingPercentage: 2_147_483_647 })
    const boundary = fixture('valid/parsed-full.json') as NativePptxDeck
    const shape = boundary.slides[0]!.elements.find((element) => element.kind === 'shape')
    const table = boundary.slides[0]!.elements.find((element) => element.kind === 'table')
    if (!shape || shape.kind !== 'shape' || !shape.stroke || !table || table.kind !== 'table' || !table.table.rows[1]![1]!.border) throw new Error('boundary fixture is incomplete')
    shape.stroke.widthEmu = PPTX_NATIVE_RESOURCE_LIMITS.maxLineWidthEmu
    shape.stroke.join = 'miter'
    shape.stroke.miterLimit = PPTX_NATIVE_RESOURCE_LIMITS.maxDrawingPercentage
    table.table.rows[1]![1]!.border!.widthEmu = PPTX_NATIVE_RESOURCE_LIMITS.maxLineWidthEmu
    expect(validateNativePptx(boundary)).toEqual({ ok: true, value: boundary })
  })

  it('requires opaque chart owner and relationship references to agree', () => {
    const result = validateNativePptx(fixture('invalid/chart-reference-mismatch.json'))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.filter((issue) => issue.code === 'native.chartReference').map((issue) => issue.path)).toEqual([
      '$.slides[0].elements[0].chart.opaqueRef.ownerPart',
      '$.slides[0].elements[0].chart.relationshipId',
    ])
  })

  it('enforces native anchor ownership and diagnostic scope ownership', () => {
    const vectors: [string, string[]][] = [
      ['chart-part-preview-type.json', ['native.assetType']],
      ['cross-slide-element-source.json', ['native.sourcePart']],
      ['parsed-element-under-authored-slide.json', ['native.sourceOwnership']],
      ['diagnostic-cross-slide-scope.json', ['native.scopeOwnership']],
    ]
    for (const [name, codes] of vectors) {
      const result = validateNativePptx(fixture(`invalid/${name}`))
      expect(result.ok, name).toBe(false)
      if (!result.ok) {
        const issueCodes = result.issues.map((issue) => issue.code)
        expect(issueCodes, name).toEqual(expect.arrayContaining(codes))
        if (name === 'chart-part-preview-type.json') expect(issueCodes).not.toContain('native.chartPart')
      }
    }
  })

  it('refuses percent-encoded traversal, separators, and controls', () => {
    for (const name of [
      'unsafe-part-encoded-parent.json', 'unsafe-part-encoded-slash.json',
      'unsafe-part-encoded-backslash.json', 'unsafe-part-encoded-control.json',
      'unsafe-part-trailing-dot.json', 'unsafe-part-encoded-trailing-dot.json',
    ]) {
      const result = validateNativePptx(fixture(`invalid/${name}`))
      expect(result.ok, name).toBe(false)
      if (!result.ok) expect(result.issues.map((issue) => issue.code), name).toContain('schema.pattern')
    }
  })

  it('requires compatibility status to aggregate opaque descendants', () => {
    const deck = fixture('valid/parsed-full.json') as NativePptxDeck
    deck.compatibility = { status: 'editable', diagnostics: [] }
    deck.slides[0].compatibility = { status: 'editable', diagnostics: [] }
    const result = validateNativePptx(deck)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.filter((issue) => issue.code === 'native.compatibilityAggregate')).toHaveLength(2)
  })

  it('canonicalizes assets without changing semantic array order', () => {
    const deck = fixture('valid/parsed-full.json') as NativePptxDeck
    const normalized = canonicalizeNativePptx(deck)
    expect(normalized.assets.map((asset) => asset.id)).toEqual(['a-preview', 'z-picture'])
    expect(normalized.slides[0].elements.map((element) => element.id)).toEqual(deck.slides[0].elements.map((element) => element.id))
    expect(deck.assets.map((asset) => asset.id)).toEqual(['z-picture', 'a-preview'])
  })

  it('matches the cross-runtime canonical golden', () => {
    const deck = fixture('valid/parsed-full.json') as NativePptxDeck
    const golden = readFileSync(resolve(fixtureRoot, 'parsed-full.canonical.json'), 'utf8')
    expect(stringifyNativePptx(deck)).toBe(golden)
  })

  it('validates before canonical stringification', () => {
    const deck = fixture('valid/authored-minimal.json') as NativePptxDeck
    delete (deck.size as Partial<NativePptxDeck['size']>).cx
    expect(() => stringifyNativePptx(deck)).toThrow(/invalid native PPTX contract/)
  })

  it('refuses group nesting beyond the shared resource depth', () => {
    const deck = fixture('valid/authored-minimal.json') as NativePptxDeck
    let element: NativeElement = {
      kind: 'connector', id: 'leaf', provenance: 'authored',
      transform: { x: 0, y: 0, cx: 1, cy: 1 }, passthrough: [],
      compatibility: { status: 'editable', diagnostics: [] },
    }
    for (let depth = 0; depth <= PPTX_NATIVE_RESOURCE_LIMITS.maxDepth; depth++) {
      element = {
        kind: 'group', id: `group-${depth}`, provenance: 'authored', children: [element],
        transform: { x: 0, y: 0, cx: 1, cy: 1 }, passthrough: [],
        compatibility: { status: 'editable', diagnostics: [] },
      }
    }
    deck.slides.push({
      id: 'slide-depth', provenance: 'authored', elements: [element], passthrough: [],
      compatibility: { status: 'editable', diagnostics: [] },
    })
    const result = validateNativePptx(deck)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.resourceDepth')
  })

  it('requires child coordinate transforms to compose to an exact affine mapping', () => {
    const deck = fixture('valid/parsed-full.json') as NativePptxDeck
    const group = deck.slides[0]!.elements.find((element) => element.kind === 'group')
    if (!group || group.kind !== 'group') throw new Error('group fixture is incomplete')
    group.childTransform = { x: 0, y: 0, cx: 3, cy: 3 }
    const result = validateNativePptx(deck)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.groupTransform')
  })

  it('requires parsed groups to retain their authoritative child coordinate transform', () => {
    const deck = fixture('valid/parsed-full.json') as NativePptxDeck
    const group = deck.slides[0]!.elements.find((element) => element.kind === 'group')
    if (!group || group.kind !== 'group') throw new Error('group fixture is incomplete')
    delete group.childTransform
    const result = validateNativePptx(deck)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.groupTransform')
  })

  it('binds native table text summaries to complete cell paragraphs and body layout', () => {
    const exact = fixture('valid/parsed-full.json') as NativePptxDeck
    const table = exact.slides[0]!.elements.find((element) => element.kind === 'table')
    if (!table || table.kind !== 'table') throw new Error('table fixture is incomplete')
    for (const row of table.table.rows) for (const cell of row) {
      delete cell.align
      cell.paragraphs = [{ align: 'left', level: 0, bullet: false, runs: [{ text: cell.text, fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }] }]
      cell.textBody = {
        leftInsetEmu: 0, rightInsetEmu: 0, topInsetEmu: 0, bottomInsetEmu: 0,
        wrap: 'square', verticalAnchor: 'top', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow',
      }
    }
    table.table.rows[0]![0] = {
      text: 'AB\nC',
      paragraphs: [
        { align: 'left', level: 0, bullet: false, runs: [{ text: 'A', fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }, { text: 'B', bold: true, fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }] },
        { align: 'right', level: 0, bullet: false, runs: [{ text: 'C', italic: true, fontFamily: 'Aptos', fontSizeHundredthPt: 1_000 }] },
      ],
      textBody: {
        leftInsetEmu: 10_000, rightInsetEmu: 20_000, topInsetEmu: 30_000, bottomInsetEmu: 40_000,
        wrap: 'square', verticalAnchor: 'top', autoFit: 'none', horizontalOverflow: 'overflow', verticalOverflow: 'overflow',
      },
      fill: 'AABBCC',
    }
    expect(validateNativePptx(exact)).toEqual({ ok: true, value: exact })

    for (const [mutate, expectedPath, expectedCode] of [
      [(deck: NativePptxDeck) => { delete tableCell(deck).textBody }, '$.slides[0].elements[4].table.rows[0][0]', 'native.tableTextAuthority'],
      [(deck: NativePptxDeck) => { tableCell(deck).text = 'forged' }, '$.slides[0].elements[4].table.rows[0][0].text', 'native.tableTextAuthority'],
      [(deck: NativePptxDeck) => { tableCell(deck).align = 'center' }, '$.slides[0].elements[4].table.rows[0][0].align', 'native.tableTextAuthority'],
      [(deck: NativePptxDeck) => { tableElement(deck).table.rowHeights = [] }, '$.slides[0].elements[4].table.rowHeights', 'native.tableGeometry'],
      [(deck: NativePptxDeck) => { tableElement(deck).transform.cx++ }, '$.slides[0].elements[4].table.columnWidths', 'native.tableGeometry'],
      [(deck: NativePptxDeck) => { const cell = tableCellAt(deck, 0, 1); delete cell.paragraphs; delete cell.textBody }, '$.slides[0].elements[4].table.rows', 'native.tableTextAuthority'],
    ] as const) {
      const invalid = structuredClone(exact)
      mutate(invalid)
      const result = validateNativePptx(invalid)
      expect(result.ok).toBe(false)
      if (!result.ok) expect(result.issues).toContainEqual(expect.objectContaining({ path: expectedPath, code: expectedCode }))
    }
  })

  it('keeps generated bindings pinned to the normative schema bytes', () => {
    const bytes = readFileSync(resolve(repositoryRoot, 'schemas/pptx-native-v1.schema.json'))
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(PPTX_NATIVE_SCHEMA_SHA256)
  })

  it('provides an asserting type guard with bounded detail', () => {
    expect(() => assertNativePptx({ contractVersion: 'pptx-native/v1' })).toThrow(/invalid native PPTX contract/)
  })
})

function tableCell(deck: NativePptxDeck) {
  return tableCellAt(deck, 0, 0)
}

function tableCellAt(deck: NativePptxDeck, row: number, column: number) {
  return tableElement(deck).table.rows[row]![column]!
}

function tableElement(deck: NativePptxDeck) {
  const table = deck.slides[0]!.elements.find((element) => element.kind === 'table')
  if (!table || table.kind !== 'table') throw new Error('table fixture is incomplete')
  return table
}
