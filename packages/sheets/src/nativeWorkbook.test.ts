import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  XLSX_NATIVE_SCHEMA_SHA256,
  decodeNativeWorkbookV1,
  nativeWorkbookStyleRawProjectionSha256V1,
  projectNativeWorkbookV1,
  validateNativeWorkbookV1,
} from './index.js'
import type { NativeWorkbookV1 } from './index.js'
import { sha256Hex } from './nativeSha256.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const fixturePath = resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v1/valid/lexical-render.json')
const fixtureSource = readFileSync(fixturePath, 'utf8').trim()
const fixture = (): NativeWorkbookV1 => JSON.parse(fixtureSource) as NativeWorkbookV1
const unicodeLengthVectors = JSON.parse(readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v1/unicode-length-vectors.json'), 'utf8')) as {
  astral: string
  accepted_sheet_ascii_count: number
  rejected_sheet_ascii_count: number
  metadata_scalar_limit: number
}

describe('native XLSX v1 Go/TypeScript binding', () => {
  it('accepts the exact canonical Go round-trip fixture', () => {
    const result = decodeNativeWorkbookV1(fixtureSource)
    expect(result).toEqual({ ok: true, value: fixture() })
    expect(XLSX_NATIVE_SCHEMA_SHA256).toMatch(/^[0-9a-f]{64}$/)
  })

  it('projects lexical values, cached formulas, durable IDs, and style provenance without coercion', () => {
    const model = projectNativeWorkbookV1(fixture())
    expect(model.document_id).toBe('workbook:fixture')
    expect(model.sheets[0].id).toBe('7')
    expect(model.sheets[0].cells[0].content).toEqual({ kind: 'literal', value: { kind: 'number', storage: 'number', lexical: '001.2300', rich: false } })
    expect(model.sheets[0].cells[1].content).toEqual({ kind: 'formula', formula: { text: 'A1*2', type: 'normal', cached: { kind: 'number', storage: 'number', lexical: '2.4600', rich: false } } })
    expect(model.sheets[0].cells[2].ooxml_type).toBe('inlineStr')
    expect(model.sheets[0].cells[3].content).toEqual({ kind: 'literal', value: { kind: 'date', storage: 'date', lexical: '2026-08-27T14:03:04Z', rich: false } })
    expect(model.sheets[0].cells[4].content).toEqual({ kind: 'blank' })
    expect(model.sheets[0].cells[5].editable).toBe(false)
    expect(model.unsupported[0]).toMatchObject({ code: 'RICH_CELL_STRING', cell_ref: 'F1' })
    expect(model.styles[1].provenance).toEqual({
      style_id: 1,
      source_revision: fixture().revision,
      source_package_sha256: fixture().source.package_sha256,
      projection: 'full',
      raw_projection_sha256: fixture().styles[1].raw_projection_sha256,
    })
  })

  it('accepts Strict contracts and preserves relocated/percent-escaped actual part spelling', () => {
    const workbook = structuredClone(fixture())
    ;(workbook.source as { dialect: string; workbook_part: string }).dialect = 'strict'
    ;(workbook.source as { workbook_part: string }).workbook_part = 'Odd/%57orkbook.xml'
    ;(workbook.sheets[0] as { part_name: string }).part_name = 'Odd/Worksheets/%53heet1.xml'
    const unsupported = workbook.unsupported[0] as { id: string; part_name?: string }
    unsupported.part_name = 'Odd/Worksheets/%53heet1.xml'
    unsupported.id = `unsupported:${createHash('sha256').update(['RICH_CELL_STRING', 'rich-text', 'sheet:7', unsupported.part_name, 'F1', ''].join('\0')).digest('hex')}`
    const result = validateNativeWorkbookV1(workbook)
    expect(result.ok).toBe(true)
    const model = projectNativeWorkbookV1(workbook)
    expect(model.source.workbook_part).toBe('Odd/%57orkbook.xml')
    expect(model.sheets[0].mutation_authority.source_part).toBe('Odd/Worksheets/%53heet1.xml')
  })

  it('rejects duplicate keys at every depth before JSON.parse loses evidence', () => {
    const duplicateRoot = fixtureSource.replace('"version":1', '"version":1,"version":1')
    const duplicateCell = fixtureSource.replace('"row":0,"column":0', '"row":0,"row":0,"column":0')
    for (const source of [duplicateRoot, duplicateCell]) {
      const result = decodeNativeWorkbookV1(source)
      expect(result.ok).toBe(false)
      if (!result.ok) expect(result.issues[0]).toMatchObject({ code: 'INVALID_JSON' })
    }
  })

  it('fails closed on unknown fields, style references, lexical coercion, and resource amplification', () => {
    const vectors: NativeWorkbookV1[] = []
    const unknown = structuredClone(fixture()) as NativeWorkbookV1 & { html?: string }; unknown.html = '<table>'; vectors.push(unknown)
    const style = structuredClone(fixture()); (style.sheets[0].cells[0] as { style_id: number }).style_id = 50; vectors.push(style)
    const lexical = structuredClone(fixture()); (lexical.sheets[0].cells[0].value as { lexical: string }).lexical = '0x10'; vectors.push(lexical)
    const missingBorder = structuredClone(fixture()); delete (missingBorder.styles[1].effective as { border?: unknown }).border; vectors.push(missingBorder)
    const borderProvenance = structuredClone(fixture()); (borderProvenance.styles[1].effective.border as { origin: string; border_id?: number }).origin = 'implicit-default'; (borderProvenance.styles[1].effective.border as { border_id?: number }).border_id = 0; vectors.push(borderProvenance)
    const sheetName = structuredClone(fixture()); (sheetName.sheets[0] as { name: string }).name = 'Unsafe/Name'; vectors.push(sheetName)
    const sheets = structuredClone(fixture()); (sheets as unknown as { sheets: unknown[] }).sheets = Array.from({ length: 1025 }, () => sheets.sheets[0]); vectors.push(sheets)
    for (const value of vectors) expect(validateNativeWorkbookV1(value).ok).toBe(false)
  })

  it('uses standards-correct scalar lengths plus explicit OOXML UTF-16 bounds', () => {
    const accepted = structuredClone(fixture())
    ;(accepted.sheets[0] as { name: string }).name = 'A'.repeat(unicodeLengthVectors.accepted_sheet_ascii_count) + unicodeLengthVectors.astral
    ;(accepted.capabilities[0] as { detail?: string }).detail = 'A'.repeat(unicodeLengthVectors.metadata_scalar_limit - 1) + unicodeLengthVectors.astral
    expect(validateNativeWorkbookV1(accepted).ok).toBe(true)

    const rejectedSheet = structuredClone(accepted)
    ;(rejectedSheet.sheets[0] as { name: string }).name = 'A'.repeat(unicodeLengthVectors.rejected_sheet_ascii_count) + unicodeLengthVectors.astral
    const rejectedMetadata = structuredClone(accepted)
    ;(rejectedMetadata.capabilities[0] as { detail?: string }).detail = 'A'.repeat(unicodeLengthVectors.metadata_scalar_limit) + unicodeLengthVectors.astral
    for (const value of [rejectedSheet, rejectedMetadata]) expect(validateNativeWorkbookV1(value).ok).toBe(false)
  })

  it('rejects locale-independent case-equivalent sheet names and boundary apostrophes', () => {
    for (const [firstName, duplicateName] of [['Lexical', 'lEXICAL'], ['Σ', 'ς'], ['İ', 'i']]) {
      const duplicate = structuredClone(fixture())
      ;(duplicate.sheets[0] as { name: string }).name = firstName
      ;(duplicate as unknown as { sheets: unknown[] }).sheets = [...duplicate.sheets, { id: '8', name: duplicateName, order: 1, state: 'visible', part_name: 'Worksheets/Sheet2.xml', rows: [], columns: [], cells: [], merged_ranges: [], editable: true }]
      const duplicateResult = validateNativeWorkbookV1(duplicate)
      expect(duplicateResult.ok).toBe(false)
      if (!duplicateResult.ok) expect(duplicateResult.issues.some((item) => item.path === '/sheets/1/name' && item.message.includes('case-insensitive'))).toBe(true)
    }
    for (const name of ["'Lexical", "Lexical'"]) {
      const workbook = structuredClone(fixture()); (workbook.sheets[0] as { name: string }).name = name
      expect(validateNativeWorkbookV1(workbook).ok).toBe(false)
    }
  })

  it('cross-binds unsupported source authority to capability, location, and editability', () => {
    const workbook = structuredClone(fixture())
    const cell = workbook.sheets[0].cells[0] as { editable: boolean }
    cell.editable = false
    const location = ['CELL_EXTENSIONS', 'extensions', 'sheet:7', 'Worksheets/Sheet1.xml', 'A1', ''].join('\0')
    ;(workbook as unknown as { unsupported: unknown[] }).unsupported = [...workbook.unsupported, {
      id: `unsupported:${createHash('sha256').update(location).digest('hex')}`,
      code: 'CELL_EXTENSIONS', capability: 'extensions', scope_id: 'sheet:7',
      part_name: 'Worksheets/Sheet1.xml', cell_ref: 'A1', preservation: 'preserve-exact', message: 'preserved extension',
    }]
    expect(validateNativeWorkbookV1(workbook).ok).toBe(true)

    const metadataOnly = structuredClone(workbook); (metadataOnly.unsupported[1] as { message: string }).message = 'different bounded explanation'
    const editable = structuredClone(workbook); (editable.sheets[0].cells[0] as { editable: boolean }).editable = true
    const wrongCapability = structuredClone(workbook); (wrongCapability.unsupported[1] as { capability: string }).capability = 'styles'
    const moved = structuredClone(workbook); (moved.unsupported[1] as { cell_ref: string }).cell_ref = 'B1'
    expect(validateNativeWorkbookV1(metadataOnly).ok).toBe(true)
    for (const value of [editable, wrongCapability, moved]) expect(validateNativeWorkbookV1(value).ok).toBe(false)
  })

  it('carries group-formula refusal over absent cells without evaluating cached values', () => {
    const workbook = structuredClone(fixture())
    const sheet = workbook.sheets[0] as { editable: boolean; refusal_code?: string }
    const cell = workbook.sheets[0].cells[1] as { editable: boolean; formula?: { text: string; type: string; ref?: string; cached?: unknown } }
    sheet.editable = false; sheet.refusal_code = 'FORMULA_GROUPS'
    cell.editable = false; cell.formula = { text: 'SUM(A1:A2)', type: 'array', ref: 'B1:B2', cached: { kind: 'number', storage: 'number', lexical: '2.4600', rich: false } }
    const unsupported = [
      ['FORMULA_ARRAY', 'formula-groups', 'sheet:7', 'Worksheets/Sheet1.xml', 'B1', ''],
      ['FORMULA_GROUP_RANGE', 'formula-groups', 'sheet:7', 'Worksheets/Sheet1.xml', '', 'B1:B2'],
    ].map(([code, capability, scope_id, part_name, cell_ref, range_ref]) => {
      const location = [code, capability, scope_id, part_name, cell_ref, range_ref].join('\0')
      return { id: `unsupported:${createHash('sha256').update(location).digest('hex')}`, code, capability, scope_id, part_name, ...(cell_ref ? { cell_ref } : { range_ref }), preservation: 'preserve-exact', message: 'formula group' }
    })
    ;(workbook as unknown as { unsupported: unknown[] }).unsupported = [...workbook.unsupported, ...unsupported]
    expect(validateNativeWorkbookV1(workbook).ok).toBe(true)
    const model = projectNativeWorkbookV1(workbook)
    expect(model.sheets[0].mutation_authority).toMatchObject({ editable: false, refusal_code: 'FORMULA_GROUPS' })
    expect(model.sheets[0].cells[1].content).toEqual({ kind: 'formula', formula: cell.formula })
    expect(model.unsupported.map((item) => item.range_ref).filter(Boolean)).toEqual(['B1:B2'])

    const noncanonical = structuredClone(workbook)
    ;(noncanonical.sheets[0].cells[1].formula as { ref?: string }).ref = '$B$1:$B$2'
    const noncanonicalRange = noncanonical.unsupported.find((item) => item.code === 'FORMULA_GROUP_RANGE') as { id: string; code: string; capability: string; scope_id: string; part_name?: string; range_ref?: string }
    noncanonicalRange.range_ref = '$B$1:$B$2'
    noncanonicalRange.id = unsupportedID(noncanonicalRange)
    const noncanonicalResult = validateNativeWorkbookV1(noncanonical)
    expect(noncanonicalResult.ok).toBe(false)
    if (!noncanonicalResult.ok) expect(noncanonicalResult.issues.some((item) => item.message.includes('canonical bounded A1'))).toBe(true)

    const overlapping = structuredClone(workbook)
    const second = overlapping.sheets[0].cells[2] as unknown as Record<string, unknown>
    delete second.value
    second.ooxml_type = 'n'; second.formula = { text: 'SUM(A1:C2)', type: 'array', ref: 'A1:C2' }; second.editable = false
    ;(overlapping.sheets[0].cells[0] as { editable: boolean }).editable = false
    const overlapItems = [
      { code: 'FORMULA_ARRAY', capability: 'formula-groups', scope_id: 'sheet:7', part_name: 'Worksheets/Sheet1.xml', cell_ref: 'C1', preservation: 'preserve-exact' as const, message: 'overlapping group' },
      { code: 'FORMULA_GROUP_RANGE', capability: 'formula-groups', scope_id: 'sheet:7', part_name: 'Worksheets/Sheet1.xml', range_ref: 'A1:C2', preservation: 'preserve-exact' as const, message: 'overlapping group' },
    ].map((item) => ({ ...item, id: unsupportedID(item) }))
    ;(overlapping as unknown as { unsupported: unknown[] }).unsupported = [...overlapping.unsupported, ...overlapItems]
    const overlapResult = validateNativeWorkbookV1(overlapping)
    expect(overlapResult.ok).toBe(false)
    if (!overlapResult.ok) expect(overlapResult.issues.map((item) => item.message)).toContain('formula group ranges must not overlap')

    const sparse = structuredClone(workbook)
    sparse.sheets[0].cells.forEach((item) => { (item as { editable: boolean }).editable = false })
    ;(sparse.sheets[0].cells[1].formula as { ref?: string }).ref = 'A1:XFD1048576'
    const sparseRange = sparse.unsupported.find((item) => item.code === 'FORMULA_GROUP_RANGE') as { id: string; code: string; capability: string; scope_id: string; part_name?: string; range_ref?: string }
    sparseRange.range_ref = 'A1:XFD1048576'; sparseRange.id = unsupportedID(sparseRange)
    expect(validateNativeWorkbookV1(sparse).ok).toBe(true)
    expect(projectNativeWorkbookV1(sparse).sheets[0].cells).toHaveLength(6)
  })

  it('does not leak a partial render model from an invalid authority contract', () => {
    const workbook = structuredClone(fixture())
    ;(workbook.sheets[0] as { editable: boolean; refusal_code?: string }).editable = false
    ;(workbook.sheets[0] as { refusal_code?: string }).refusal_code = 'SHEET_PROTECTION'
    expect(() => projectNativeWorkbookV1(workbook)).toThrow(/sheet refusal does not match authoritative unsupported source/)
  })
})

describe('native XLSX authority SHA-256', () => {
  it('matches Node crypto for standard, multi-block, and non-BMP UTF-8 vectors', () => {
    for (const value of ['', 'abc', 'a'.repeat(1000), '😀����✨']) {
      expect(sha256Hex(value)).toBe(createHash('sha256').update(value, 'utf8').digest('hex'))
    }
  })

  it('matches Go compact JSON escaping for canonical style projection digests', () => {
    expect(nativeWorkbookStyleRawProjectionSha256V1({
      font_name: 'A\u2028B\u2029C', fill: { origin: 'implicit-default' }, border: { origin: 'implicit-default' }, projection: 'full', unsupported: [],
    })).toBe('sha256:29a0ff04e8fc39c1f1844e6d23e1e6411c500577f7e91dd90c1b497afd276098')
  })
})

function unsupportedID(item: { code: string; capability: string; scope_id: string; part_name?: string; cell_ref?: string; range_ref?: string }): string {
  const location = [item.code, item.capability, item.scope_id, item.part_name ?? '', item.cell_ref ?? '', item.range_ref ?? ''].join('\0')
  return `unsupported:${createHash('sha256').update(location).digest('hex')}`
}

describe('native XLSX runtime dependency boundary', () => {
  it('uses only local ECMAScript modules and contains no DOM/HTML/PDF/Univer/browser API', () => {
    const runtimeFiles = ['nativeJson.ts', 'nativeSchemaValidation.ts', 'nativeValidation.ts', 'nativeRenderModel.ts', 'nativeSheetGeometry.ts', 'nativeSheetDecorations.ts', 'nativeSha256.ts', 'nativeContract.generated.ts']
    for (const file of runtimeFiles) {
      const source = readFileSync(resolve(import.meta.dirname, file), 'utf8')
      const imports = [...source.matchAll(/from\s+['"]([^'"]+)['"]/g)].map((match) => match[1])
      expect(imports.every((specifier) => specifier.startsWith('./')), `${file} imports ${imports.join(', ')}`).toBe(true)
      expect(source, file).not.toMatch(/\b(?:window|document|HTMLElement|HTMLCanvasElement|WorkerGlobalScope|PDFDocument|Electron|Univer|Konva|React)\b/)
      expect(imports, file).not.toEqual(expect.arrayContaining([expect.stringMatching(/^(?:@injoffice\/pdf|pdfjs|pdf-lib)(?:\/|$)/)]))
      expect(source, file).not.toMatch(/<\/?(?:html|body|table|div|span)\b/i)
    }
  })
})
