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
  nativeNormalFontDescentEmV1,
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

/** Offset of the named table's 16-byte directory record in the fixture. */
function tableRecordOffset(tag: string): number {
  const count = FONT_BYTES[4]! * 256 + FONT_BYTES[5]!
  for (let index = 0; index < count; index++) {
    const at = 12 + index * 16
    if (String.fromCharCode(FONT_BYTES[at]!, FONT_BYTES[at + 1]!, FONT_BYTES[at + 2]!, FONT_BYTES[at + 3]!) === tag) return at
  }
  throw new Error(`fixture has no ${tag} table`)
}

/** The fixture with one 32-bit field of one table record overwritten. */
function editTableRecord(tag: string, fieldOffset: number, value: number): Uint8Array {
  const bytes = new Uint8Array(FONT_BYTES)
  const at = tableRecordOffset(tag) + fieldOffset
  bytes[at] = (value >>> 24) & 0xff
  bytes[at + 1] = (value >>> 16) & 0xff
  bytes[at + 2] = (value >>> 8) & 0xff
  bytes[at + 3] = value & 0xff
  return bytes
}

/** The fixture with one table's STORED checksum made stale; no table byte moves. */
function staleChecksumFont(tag: string): Uint8Array {
  const at = tableRecordOffset(tag) + 4
  const stored = ((FONT_BYTES[at]! * 0x1000000) + (FONT_BYTES[at + 1]! << 16) + (FONT_BYTES[at + 2]! << 8) + FONT_BYTES[at + 3]!) >>> 0
  return editTableRecord(tag, 4, (stored ^ 0xa5a5a5a5) >>> 0)
}

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
    // Whole POINTS, on the page's own rounded maximum digit width (7 pt for
    // this fixture's DejaVu Sans Normal), not 96-dpi CSS pixels on its 9-px
    // screen one: the grid lattice belongs to the device it is printed on.
    expect(geometry.columns).toEqual([
      { column: 0, x_emu: 0, width_emu: 87 * EMU_PER_POINT, hidden: false, width_characters: 12.5, source: 'column-override' },
      { column: 1, x_emu: 1_104_900, width_emu: 87 * EMU_PER_POINT, hidden: false, width_characters: 12.5, source: 'column-override' },
      { column: 2, x_emu: 2_209_800, width_emu: 61 * EMU_PER_POINT, hidden: false, width_characters: 8.7109375, source: 'sheet-default' },
    ])
    expect(geometry.bounds).toEqual({ x_emu: 0, y_emu: 0, width_emu: 2_984_500, height_emu: 612_775 })
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

  /**
   * Excel rounds the maximum digit width to a whole unit of the device it is
   * laying out on, then floors every column to a whole unit of that same
   * device (ECMA-376 §18.3.1.13). The printed page's device is the page: 72
   * units per inch. So the printed grid and the 96-dpi screen grid genuinely
   * disagree, and page breaks follow the printed one.
   *
   * DejaVu Sans at 8.75 pt has exactly Calibri 11's split - 7 CSS pixels on a
   * 96-dpi screen, 6 points on the page - so this fixture reproduces the
   * measurement without needing Calibri's bytes: a default `baseColWidth` 8
   * column is 61 px of screen grid and 53 pt of printed grid, and Excel
   * 16.112.4 prints 53.
   */
  it('builds the grid on the page device maximum digit width, not the 96-dpi screen one', () => {
    const workbook = structuredClone(fixture())
    ;(workbook as unknown as { normal_style: { font_size_points: number } }).normal_style.font_size_points = 8.75
    const authority = metric(workbook)
    expect(authority.measurement_dpi).toBe(96)
    expect(authority.maximum_digit_width_pixels).toBe(7)
    expect(authority.page_measurement_dpi).toBe(72)
    expect(authority.page_maximum_digit_width_points).toBe(6)

    const geometry = compileNativeSheetGeometryV2(projectNativeWorkbookV2(workbook), '7', { row: 0, column: 0, end_row: 0, end_column: 2 }, authority)
    const sheetDefault = geometry.columns[2]!
    expect(sheetDefault.source).toBe('sheet-default')
    // Padded on the page metric (8 + 5/6), not on the screen metric (8 + 5/7).
    expect(sheetDefault.width_characters).toBe(paddedBaseColumnWidth(8, 6))
    expect(sheetDefault.width_emu).toBe(53 * EMU_PER_POINT)
    expect(sheetDefault.width_emu).not.toBe(61 * EMU_PER_CSS_PIXEL)
    // A stored width travels the same lattice on the same device.
    expect(geometry.columns[0]!.width_emu).toBe(characterWidthToPixels(12.5, 6) * EMU_PER_POINT)
  })

  /**
   * Reference guard, NOT a change detector: `characterWidthToPixels` and
   * `paddedBaseColumnWidth` are untouched by this commit and every line below
   * passes without it. They pin the Excel 16.112.4 numbers the device choice
   * above was derived from, each read off that workbook's own PDF export as a
   * page clip or cell fill rectangle rather than inferred: `pivot_dark1`
   * (Calibri 11, no `<col>`: 53 pt columns, page clip 500 = 6x53 + 84 + 98),
   * `databar` (Arial 10 at width="11.5204081632653": 69 pt, page clip 483 =
   * 7x69), `tdf130104_indent` (Times New Roman 10 at 25.5 and 30.6640625: page
   * clip 280 = 127 + 153) and `autofilter-colors` (Liberation Sans 11 at
   * width="10.625": page clip 192 = 3x64). The last line is the same default
   * column on the screen device, which is what we used to print.
   */
  it('reproduces the Excel-measured printed column widths from the page-device metric', () => {
    expect(characterWidthToPixels(paddedBaseColumnWidth(8, 6), 6)).toBe(53)
    expect(characterWidthToPixels(11.5204081632653, 6)).toBe(69)
    expect(characterWidthToPixels(25.5, 5)).toBe(127)
    expect(characterWidthToPixels(30.6640625, 5)).toBe(153)
    expect(characterWidthToPixels(10.625, 6)).toBe(64)
    expect(characterWidthToPixels(paddedBaseColumnWidth(8, 7), 7)).toBe(61)
  })

  it('preserves hidden-axis coordinates and produces a merged rectangle without expanding sparse cells', () => {
    const workbook = workbookWithMerge()
    const geometry = compileNativeSheetGeometryV2(projectNativeWorkbookV2(workbook), '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, metric(workbook))
    expect(geometry.rows[1]).toMatchObject({ row: 1, y_emu: 231_775, height_emu: 0, hidden: true })
    expect(geometry.merged_ranges).toEqual([{
      ref: 'B2:C3', row: 1, column: 1, end_row: 2, end_column: 2,
      rect: { x_emu: 1_104_900, y_emu: 231_775, width_emu: 1_879_600, height_emu: 190_500 },
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

  // sheetFormatPr is optional in CT_Worksheet (ECMA-376 §18.3.1.99), and the
  // LibreOffice-written formats.xlsx omits it while stating ht on every row it
  // has. Nothing is inferred for such a sheet: the stored heights are compiled
  // and the absent default is never read.
  it('compiles a worksheet with no sheetFormatPr from the heights its rows state', () => {
    const workbook = structuredClone(fixture())
    delete (workbook.sheets[0] as { sheet_format?: unknown }).sheet_format
    const geometry = compileNativeSheetGeometryV2(projectNativeWorkbookV2(workbook), '7', { row: 0, column: 0, end_row: 0, end_column: 2 }, metric(workbook))

    // The numbers are the ones the source states, identical to the same sheet
    // compiled with a sheetFormatPr present.
    const withFormat = compileNativeSheetGeometryV2(projectNativeWorkbookV2(fixture()), '7', { row: 0, column: 0, end_row: 0, end_column: 2 }, metric(fixture()))
    expect(geometry.rows).toEqual(withFormat.rows)
    expect(geometry.columns).toEqual(withFormat.columns)
    expect(geometry.rows[0]!.height_emu).toBe(Math.round(18.25 * EMU_PER_POINT))

    // The third column states no width of its own; baseColWidth does have a
    // declared default of 8 (§18.3.1.81), so it is padded from that either way.
    expect(geometry.columns[2]!.source).toBe('sheet-default')
    expect(geometry.columns[2]!.width_characters).toBe(paddedBaseColumnWidth(8, metric(workbook).page_maximum_digit_width_points))
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
      // Row 1 states no ht of its own, so with no sheetFormatPr there is no
      // default row height to read and the sheet still refuses.
      () => compileNativeSheetGeometryV2(projectNativeWorkbookV2(withoutFormat), '7', { row: 0, column: 0, end_row: 1, end_column: 0 }, metric(withoutFormat)),
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

describe('normal font descent', () => {
  it('reads the hhea descent as a fraction of the em square from the same qualified font bytes', () => {
    // DejaVu Sans states unitsPerEm 2048 and hhea descender -483. Excel places a
    // bottom-aligned baseline exactly this far above the row's bottom edge, so
    // the preview needs the face's own number rather than a constant inset.
    expect(nativeNormalFontDescentEmV1(FONT_BYTES)).toBeCloseTo(483 / 2048, 12)
  })

  it('refuses font bytes it cannot qualify rather than guessing a descent', () => {
    expect(() => nativeNormalFontDescentEmV1(new Uint8Array(8))).toThrow(NativeSheetGeometryV2Error)
    expect(() => nativeNormalFontDescentEmV1(FONT_BYTES.subarray(0, 4096))).toThrow(NativeSheetGeometryV2Error)
    expect(() => nativeNormalFontDescentEmV1(new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf'))).fill(0, 0, 4))).toThrow(NativeSheetGeometryV2Error)
  })

  // A stale table checksum is advisory build metadata, not a parse
  // precondition. Word 16's own shipped Symbol carries one on its `cmap`, and
  // refusing it cost the whole workbook preview an error about a font table.
  it('reads a face whose stored table checksum is stale exactly as it reads the pristine face', () => {
    const stale = staleChecksumFont('cmap')
    // Only the four stored checksum bytes differ; every table byte is identical.
    expect(stale.length).toBe(FONT_BYTES.length)
    expect(stale.reduce((count, byte, index) => byte === FONT_BYTES[index] ? count : count + 1, 0)).toBe(4)
    expect(nativeNormalFontDescentEmV1(stale)).toBe(nativeNormalFontDescentEmV1(FONT_BYTES))
    const pristine = metric(fixture())
    const authority = createNativeMaximumDigitWidthAuthorityV2(projectNativeWorkbookV2(fixture()), stale)
    expect(authority.maximum_digit_width_pixels).toBe(pristine.maximum_digit_width_pixels)
    // Every table's record is advisory, `head` included.
    expect(nativeNormalFontDescentEmV1(staleChecksumFont('head'))).toBe(nativeNormalFontDescentEmV1(FONT_BYTES))
    expect(nativeNormalFontDescentEmV1(staleChecksumFont('hhea'))).toBe(nativeNormalFontDescentEmV1(FONT_BYTES))
  })

  // Inertness guard: these are the structural preconditions the checksum was
  // never what caught, and each already refused before it was dropped.
  it('keeps refusing sfnt structure the checksum was never what caught', () => {
    // A table record whose four-byte-aligned offset lies outside the font.
    expect(() => nativeNormalFontDescentEmV1(editTableRecord('cmap', 8, FONT_BYTES.length + 4))).toThrow(NativeSheetGeometryV2Error)
    // A zero-length table record.
    expect(() => nativeNormalFontDescentEmV1(editTableRecord('cmap', 12, 0))).toThrow(NativeSheetGeometryV2Error)
    // A misaligned table offset.
    expect(() => nativeNormalFontDescentEmV1(editTableRecord('cmap', 8, 13))).toThrow(NativeSheetGeometryV2Error)
    // A collection, which is not a standalone fixed TrueType sfnt.
    expect(() => nativeNormalFontDescentEmV1(new Uint8Array(FONT_BYTES).fill(0x74, 0, 4))).toThrow(NativeSheetGeometryV2Error)
  })
})
