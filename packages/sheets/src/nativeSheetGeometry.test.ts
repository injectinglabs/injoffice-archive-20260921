import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  EMU_PER_CSS_PIXEL,
  EMU_PER_POINT,
  NativeSheetGeometryError,
  characterWidthToPixels,
  compileNativeSheetGeometryV1,
  createNativeSheetGeometryRecordingSurfaceV1,
  emitNativeSheetGeometryCommandsV1,
  paddedBaseColumnWidth,
  projectNativeWorkbookV1,
  replayNativeSheetGeometryCommandsV1,
} from './index.js'
import type { NativeMaximumDigitWidthAuthorityV1, NativeWorkbookV1 } from './index.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const fixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v1/valid/lexical-render.json'), 'utf8').trim()
const fixture = (): NativeWorkbookV1 => ({
  ...(JSON.parse(fixtureSource) as NativeWorkbookV1),
  normal_style: {
    style_xf_id: 0, font_id: 0, font_name: 'Calibri', font_size_points: 11, font_bold: false, font_italic: false,
    font_record_sha256: `sha256:${'e'.repeat(64)}`,
  },
})
const metric = (workbook: NativeWorkbookV1 = fixture(), overrides: Partial<NativeMaximumDigitWidthAuthorityV1> = {}): NativeMaximumDigitWidthAuthorityV1 => ({
  source_revision: workbook.revision,
  source_package_sha256: workbook.source.package_sha256,
  normal_style_xf_id: workbook.normal_style!.style_xf_id,
  normal_style_font_id: workbook.normal_style!.font_id,
  font_name: workbook.normal_style!.font_name,
  font_size_points: workbook.normal_style!.font_size_points,
  font_bold: workbook.normal_style!.font_bold,
  font_italic: workbook.normal_style!.font_italic,
  normal_font_record_sha256: workbook.normal_style!.font_record_sha256 as `sha256:${string}`,
  font_sha256: `sha256:${'b'.repeat(64)}`,
  provider_id: 'fixture-opentype-parser',
  provider_revision: '1',
  measurement_dpi: 96,
  maximum_digit_width_pixels: 7,
  ...overrides,
})

function workbookWithMerge(): NativeWorkbookV1 {
  const workbook = structuredClone(fixture())
  const sheet = workbook.sheets[0] as unknown as { rows: unknown[]; merged_ranges: unknown[] }
  sheet.rows = [
    ...sheet.rows,
    { row: 1, height_points: 20, hidden: true, custom_height: true },
  ]
  sheet.merged_ranges = [{ ref: 'B2:C3', row: 1, column: 1, end_row: 2, end_column: 2, editable: false }]
  const location = ['MERGED_CELLS', 'merges', 'sheet:7', 'Worksheets/Sheet1.xml', '', ''].join('\0')
  ;(workbook as unknown as { unsupported: unknown[] }).unsupported = [...workbook.unsupported, {
    id: `unsupported:${createHash('sha256').update(location).digest('hex')}`,
    code: 'MERGED_CELLS', capability: 'merges', scope_id: 'sheet:7', part_name: 'Worksheets/Sheet1.xml',
    preservation: 'preserve-exact', message: 'merged source geometry',
  }]
  return workbook
}

describe('native XLSX deterministic sheet geometry', () => {
  it('compiles the shared native fixture into exact integer-EMU axis bands', () => {
    const workbook = fixture()
    const model = projectNativeWorkbookV1(workbook)
    const before = JSON.stringify(model)
    const geometry = compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))

    expect(geometry.rows).toEqual([
      { row: 0, y_emu: 0, height_emu: Math.round(18.25 * EMU_PER_POINT), hidden: false, source: 'row-override' },
      { row: 1, y_emu: 231_775, height_emu: 15 * EMU_PER_POINT, hidden: false, source: 'sheet-default' },
      { row: 2, y_emu: 422_275, height_emu: 15 * EMU_PER_POINT, hidden: false, source: 'sheet-default' },
    ])
    expect(geometry.columns).toEqual([
      { column: 0, x_emu: 0, width_emu: 87 * EMU_PER_CSS_PIXEL, hidden: false, width_characters: 12.5, source: 'column-override' },
      { column: 1, x_emu: 828_675, width_emu: 87 * EMU_PER_CSS_PIXEL, hidden: false, width_characters: 12.5, source: 'column-override' },
      { column: 2, x_emu: 1_657_350, width_emu: 61 * EMU_PER_CSS_PIXEL, hidden: false, width_characters: 8.7109375, source: 'sheet-default' },
    ])
    expect(geometry.bounds).toEqual({ x_emu: 0, y_emu: 0, width_emu: 2_238_375, height_emu: 612_775 })
    expect(geometry.geometry_sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(Object.isFrozen(geometry)).toBe(true)
    expect(Object.isFrozen(geometry.rows[0])).toBe(true)
    expect(JSON.stringify(model)).toBe(before)
  })

  it('uses the workbook MDW for Excel-compatible width vectors', () => {
    expect(characterWidthToPixels(10.6640625, 7)).toBe(75)
    expect(characterWidthToPixels(34.83203125, 8)).toBe(279)
    expect(paddedBaseColumnWidth(8, 7)).toBe(8.7109375)
    expect(paddedBaseColumnWidth(8, 8)).toBe(8.625)
    expect(characterWidthToPixels(paddedBaseColumnWidth(8, 8), 8)).toBe(69)
    expect(paddedBaseColumnWidth(10, 8)).toBe(10.625)
  })

  it('preserves hidden-axis coordinates and produces a merged rectangle without expanding sparse cells', () => {
    const workbook = workbookWithMerge()
    const geometry = compileNativeSheetGeometryV1(projectNativeWorkbookV1(workbook), '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))
    expect(geometry.rows[1]).toMatchObject({ row: 1, y_emu: 231_775, height_emu: 0, hidden: true })
    expect(geometry.merged_ranges).toEqual([{
      ref: 'B2:C3', row: 1, column: 1, end_row: 2, end_column: 2,
      rect: { x_emu: 828_675, y_emu: 231_775, width_emu: 1_409_700, height_emu: 190_500 },
    }])
    expect(geometry.rows).toHaveLength(3)
    expect(geometry.columns).toHaveLength(3)
  })

  it('is byte-deterministic across repeated compilation and renderer-neutral command recording', () => {
    const workbook = fixture()
    const model = projectNativeWorkbookV1(workbook)
    const first = compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))
    const second = compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.geometry_sha256).toBe(second.geometry_sha256)

    const reorderedMetric = {
      maximum_digit_width_pixels: 7,
      provider_revision: '1',
      provider_id: 'fixture-opentype-parser',
      measurement_dpi: 96 as const,
      font_sha256: `sha256:${'b'.repeat(64)}` as const,
      normal_font_record_sha256: `sha256:${'e'.repeat(64)}` as const,
      font_italic: false,
      font_bold: false,
      font_size_points: 11,
      font_name: 'Calibri',
      normal_style_font_id: 0,
      normal_style_xf_id: 0,
      source_package_sha256: workbook.source.package_sha256,
      source_revision: workbook.revision,
    }
    const reordered = compileNativeSheetGeometryV1(model, '7', { end_column: 2, end_row: 2, column: 0, row: 0 }, reorderedMetric)
    expect(JSON.stringify(reordered)).toBe(JSON.stringify(first))

    const recording = createNativeSheetGeometryRecordingSurfaceV1()
    emitNativeSheetGeometryCommandsV1(first, recording)
    const commands = recording.finish()
    expect(commands.map((command) => command.kind)).toEqual(['beginSheet', 'clipRect', 'rowBand', 'rowBand', 'rowBand', 'columnBand', 'columnBand', 'columnBand', 'endSheet'])
    expect(Object.isFrozen(commands)).toBe(true)
    const replayed: string[] = []
    replayNativeSheetGeometryCommandsV1(replayed, commands, { execute(target, command) { target.push(command.kind) } })
    expect(replayed).toEqual(commands.map((command) => command.kind))
  })

  it('makes nonzero viewport coordinates explicitly viewport-local', () => {
    const workbook = fixture()
    const geometry = compileNativeSheetGeometryV1(projectNativeWorkbookV1(workbook), '7', { row: 1, column: 2, end_row: 2, end_column: 2 }, metric(workbook))
    expect(geometry.coordinate_space).toBe('viewport-local')
    expect(geometry.origin_cell).toEqual({ row: 1, column: 2 })
    expect(geometry.rows[0]).toMatchObject({ row: 1, y_emu: 0 })
    expect(geometry.columns[0]).toMatchObject({ column: 2, x_emu: 0 })
  })

  it('fails closed on absent native defaults, unbound font metrics, zeroHeight, clipped merges, and budgets', () => {
    const base = fixture()
    const withoutFormat = structuredClone(base); delete (withoutFormat.sheets[0] as { sheet_format?: unknown }).sheet_format
    const zeroHeight = structuredClone(base); (zeroHeight.sheets[0].sheet_format as { zero_height: boolean }).zero_height = true
    const bestFit = structuredClone(base)
    const bestFitColumn = bestFit.sheets[0].columns[0] as { width?: number; custom_width: boolean; best_fit: boolean }
    delete bestFitColumn.width
    bestFitColumn.custom_width = false
    bestFitColumn.best_fit = true
    const merged = workbookWithMerge()
    const missingNormal = structuredClone(base); delete (missingNormal as { normal_style?: unknown }).normal_style
    const dimensionExtras = structuredClone(base)
    const dimensionLocation = ['ROW_DIMENSION_EXTRAS', 'dimensions', 'sheet:7', 'Worksheets/Sheet1.xml', '', ''].join('\0')
    ;(dimensionExtras as unknown as { unsupported: unknown[] }).unsupported = [...dimensionExtras.unsupported, {
      id: `unsupported:${createHash('sha256').update(dimensionLocation).digest('hex')}`, code: 'ROW_DIMENSION_EXTRAS', capability: 'dimensions', scope_id: 'sheet:7',
      part_name: 'Worksheets/Sheet1.xml', preservation: 'preserve-exact', message: 'thickTop remains source-authoritative',
    }]
    const cases: Array<() => unknown> = [
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(base), '7', { row: -0, column: 0, end_row: 0, end_column: 0 }, metric(base)),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(withoutFormat), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(withoutFormat)),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(base), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(base, { source_revision: `rev:${'c'.repeat(64)}` })),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(base), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(base, { measurement_dpi: 192 as 96 })),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(zeroHeight), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(zeroHeight)),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(bestFit), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(bestFit)),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(missingNormal), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(base)),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(dimensionExtras), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(dimensionExtras)),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(merged), '7', { row: 1, column: 1, end_row: 1, end_column: 2 }, metric(merged)),
      () => compileNativeSheetGeometryV1(projectNativeWorkbookV1(base), '7', { row: 0, column: 0, end_row: 999, end_column: 999 }, metric(base)),
    ]
    for (const run of cases) expect(run).toThrow(NativeSheetGeometryError)

    const geometry = compileNativeSheetGeometryV1(projectNativeWorkbookV1(base), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(base))
    const recording = createNativeSheetGeometryRecordingSurfaceV1(2)
    expect(() => emitNativeSheetGeometryCommandsV1(geometry, recording)).toThrowError(expect.objectContaining({ code: 'geometry.commandBudget' }))
    expect(recording.commands).toEqual([])
  })
})
