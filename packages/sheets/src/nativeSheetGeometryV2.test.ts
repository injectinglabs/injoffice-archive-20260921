import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import {
  EMU_PER_CSS_PIXEL,
  EMU_PER_POINT,
  NativeSheetGeometryV2Error,
  characterWidthToPixels,
  compileNativeSheetGeometryV2,
  createNativeMaximumDigitWidthAuthorityV2,
  createNativeSheetGeometryRecordingSurfaceV2,
  emitNativeSheetGeometryCommandsV2,
  paddedBaseColumnWidth,
  projectNativeWorkbookV2,
  replayNativeSheetGeometryCommandsV2,
  validateNativeSheetGeometryV2,
} from './index.js'
import type { NativeMaximumDigitWidthAuthorityV2, NativeWorkbookV2 } from './index.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const require = createRequire(import.meta.url)
const FONT_BYTES = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const fixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json'), 'utf8').trim()
const fixture = (): NativeWorkbookV2 => ({
  ...(JSON.parse(fixtureSource) as NativeWorkbookV2),
  normal_style: {
    style_xf_id: 0, font_id: 0, font_name: 'DejaVu Sans', font_size_points: 11, font_bold: false, font_italic: false,
    font_record_sha256: `sha256:${'e'.repeat(64)}`,
  },
})
const metric = (workbook: NativeWorkbookV2): NativeMaximumDigitWidthAuthorityV2 => createNativeMaximumDigitWidthAuthorityV2(projectNativeWorkbookV2(workbook), FONT_BYTES)

function workbookWithMerge(): NativeWorkbookV2 {
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
    const model = projectNativeWorkbookV2(workbook)
    const before = JSON.stringify(model)
    const geometry = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))

    expect(geometry.rows).toEqual([
      { row: 0, y_emu: 0, height_emu: Math.round(18.25 * EMU_PER_POINT), hidden: false, source: 'row-override' },
      { row: 1, y_emu: 231_775, height_emu: 15 * EMU_PER_POINT, hidden: false, source: 'sheet-default' },
      { row: 2, y_emu: 422_275, height_emu: 15 * EMU_PER_POINT, hidden: false, source: 'sheet-default' },
    ])
    expect(geometry.columns).toEqual([
      { column: 0, x_emu: 0, width_emu: 112 * EMU_PER_CSS_PIXEL, hidden: false, width_characters: 12.5, source: 'column-override' },
      { column: 1, x_emu: 1_066_800, width_emu: 112 * EMU_PER_CSS_PIXEL, hidden: false, width_characters: 12.5, source: 'column-override' },
      { column: 2, x_emu: 2_133_600, width_emu: 77 * EMU_PER_CSS_PIXEL, hidden: false, width_characters: 8.5546875, source: 'sheet-default' },
    ])
    expect(geometry.bounds).toEqual({ x_emu: 0, y_emu: 0, width_emu: 2_867_025, height_emu: 612_775 })
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
    const geometry = compileNativeSheetGeometryV2(projectNativeWorkbookV2(workbook), '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))
    expect(geometry.rows[1]).toMatchObject({ row: 1, y_emu: 231_775, height_emu: 0, hidden: true })
    expect(geometry.merged_ranges).toEqual([{
      ref: 'B2:C3', row: 1, column: 1, end_row: 2, end_column: 2,
      rect: { x_emu: 1_066_800, y_emu: 231_775, width_emu: 1_800_225, height_emu: 190_500 },
    }])
    expect(geometry.rows).toHaveLength(3)
    expect(geometry.columns).toHaveLength(3)
  })

  it('is byte-deterministic across repeated compilation and renderer-neutral command recording', () => {
    const workbook = fixture()
    const model = projectNativeWorkbookV2(workbook)
    const first = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))
    const second = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.geometry_sha256).toBe(second.geometry_sha256)

    const reordered = compileNativeSheetGeometryV2(model, '7', { end_column: 2, end_row: 2, column: 0, row: 0 }, metric(workbook))
    expect(JSON.stringify(reordered)).toBe(JSON.stringify(first))

    const recording = createNativeSheetGeometryRecordingSurfaceV2()
    emitNativeSheetGeometryCommandsV2(first, recording)
    const commands = recording.finish()
    expect(commands.map((command) => command.kind)).toEqual(['beginSheet', 'clipRect', 'rowBand', 'rowBand', 'rowBand', 'columnBand', 'columnBand', 'columnBand', 'endSheet'])
    expect(Object.isFrozen(commands)).toBe(true)
    const replayed: string[] = []
    replayNativeSheetGeometryCommandsV2(replayed, commands, { execute(target, command) { target.push(command.kind) } })
    expect(replayed).toEqual(commands.map((command) => command.kind))
  })

  it('validates geometry plans and complete command streams before any host callback', () => {
    const workbook = fixture()
    const model = projectNativeWorkbookV2(workbook)
    const geometry = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 1, end_column: 1 }, metric(workbook))
    expect(JSON.stringify(validateNativeSheetGeometryV2(geometry))).toBe(JSON.stringify(geometry))

    const stale = structuredClone(geometry)
    ;(stale as { geometry_sha256: string }).geometry_sha256 = `sha256:${'0'.repeat(64)}`
    let pushes = 0
    expect(() => emitNativeSheetGeometryCommandsV2(stale, { push() { pushes++ } })).toThrowError(expect.objectContaining({ code: 'geometry.planDigest' }))
    expect(pushes).toBe(0)

    const invalid = structuredClone(geometry)
    ;(invalid.rows[1] as { y_emu: number }).y_emu++
    const { geometry_sha256: _oldDigest, ...unsigned } = invalid
    ;(invalid as { geometry_sha256: string }).geometry_sha256 = `sha256:${createHash('sha256').update(`injoffice.xlsx.sheet-geometry.v1\0${JSON.stringify(unsigned)}`).digest('hex')}`
    pushes = 0
    expect(() => emitNativeSheetGeometryCommandsV2(invalid, { push() { pushes++ } })).toThrowError(expect.objectContaining({ code: 'geometry.planInvalid' }))
    expect(pushes).toBe(0)

    const recording = createNativeSheetGeometryRecordingSurfaceV2()
    emitNativeSheetGeometryCommandsV2(geometry, recording)
    const commands = structuredClone(recording.finish())
    const row = commands.find((command) => command.kind === 'rowBand')!
    ;(row.band as { height_emu: number }).height_emu++
    let callbacks = 0
    expect(() => replayNativeSheetGeometryCommandsV2({}, commands, { execute() { callbacks++ } })).toThrow(NativeSheetGeometryV2Error)
    expect(callbacks).toBe(0)

    let accessorCalls = 0
    const accessorSurface = {}
    Object.defineProperty(accessorSurface, 'push', { enumerable: true, get() { accessorCalls++; return () => undefined } })
    expect(() => emitNativeSheetGeometryCommandsV2(geometry, accessorSurface as never)).toThrowError(expect.objectContaining({ code: 'geometry.commandInvalid' }))
    expect(accessorCalls).toBe(0)
  })

  it('refuses accessor, Proxy, negative-zero viewport, and caller-labeled metric inputs', () => {
    const workbook = fixture(), model = projectNativeWorkbookV2(workbook), authority = metric(workbook)
    let getterCalls = 0
    const accessorViewport = { column: 0, end_row: 0, end_column: 0 }
    Object.defineProperty(accessorViewport, 'row', { enumerable: true, get() { getterCalls++; return 0 } })
    expect(() => compileNativeSheetGeometryV2(model, '7', accessorViewport as never, authority)).toThrow(NativeSheetGeometryV2Error)
    expect(getterCalls).toBe(0)
    let reads = 0
    const proxiedViewport = new Proxy({ row: 0, column: 0, end_row: 0, end_column: 0 }, { get(target, property, receiver) { reads++; return Reflect.get(target, property, receiver) } })
    expect(() => compileNativeSheetGeometryV2(model, '7', proxiedViewport, authority)).toThrow(NativeSheetGeometryV2Error)
    expect(reads).toBe(0)
    expect(() => compileNativeSheetGeometryV2(model, '7', { row: -0, column: 0, end_row: 0, end_column: 0 }, authority)).toThrow(NativeSheetGeometryV2Error)
    expect(() => compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, { ...authority })).toThrowError(expect.objectContaining({ code: 'geometry.metricAuthority' }))
  })

  it('makes nonzero viewport coordinates explicitly viewport-local', () => {
    const workbook = fixture()
    const geometry = compileNativeSheetGeometryV2(projectNativeWorkbookV2(workbook), '7', { row: 1, column: 2, end_row: 2, end_column: 2 }, metric(workbook))
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
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(withoutFormat), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(withoutFormat)),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(base), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, { ...metric(base), source_revision: `rev:${'c'.repeat(64)}` }),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(base), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, { ...metric(base), measurement_dpi: 192 as 96 }),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(zeroHeight), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(zeroHeight)),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(bestFit), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(bestFit)),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(missingNormal), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(base)),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(dimensionExtras), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(dimensionExtras)),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(merged), '7', { row: 1, column: 1, end_row: 1, end_column: 2 }, metric(merged)),
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(base), '7', { row: 0, column: 0, end_row: 999, end_column: 999 }, metric(base)),
    ]
    for (const run of cases) expect(run).toThrow(NativeSheetGeometryV2Error)

    const geometry = compileNativeSheetGeometryV2(projectNativeWorkbookV2(base), '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(base))
    const recording = createNativeSheetGeometryRecordingSurfaceV2(2)
    expect(() => emitNativeSheetGeometryCommandsV2(geometry, recording)).toThrowError(expect.objectContaining({ code: 'geometry.commandBudget' }))
    expect(recording.commands).toEqual([])
  })
})
