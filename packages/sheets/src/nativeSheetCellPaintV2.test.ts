import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import {
  NativeSheetCellPaintError,
  NativeWorkbookV2ValidationError,
  compileNativeSheetCellPaintV2,
  compileNativeSheetGeometryV2,
  createNativeMaximumDigitWidthAuthorityV2,
  createNativeSheetCellPaintRecordingSurfaceV2,
  emitNativeSheetCellPaintCommandsV2,
  formatNativeSheetCellDisplayV2,
  nativeWorkbookStyleRawProjectionSha256V2,
  projectNativeWorkbookV2,
  replayNativeSheetCellPaintCommandsV2,
  validateNativeSheetCellPaintPlanV2,
} from './index.js'
import type {
  NativeMaximumDigitWidthAuthorityV2,
  NativeSheetCellPaintCommandV2,
  NativeSheetCellPaintPlanV2,
  NativeSheetGeometryV2,
  NativeWorkbookRenderModelV2,
  NativeWorkbookUnsupportedV2,
  NativeWorkbookV2,
} from './index.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const require = createRequire(import.meta.url)
const FONT_BYTES = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const fixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json'), 'utf8').trim()
const excelFixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json'), 'utf8').trim()

type MutableWorkbook = {
  date1904?: boolean
  normal_style?: NativeWorkbookV2['normal_style']
  styles: Array<{ id: number; effective: MutableEffective; raw_projection_sha256: string }>
  sheets: Array<{
    id: string; part_name: string; editable: boolean; refusal_code?: string
    rows: Array<{ row: number; height_points?: number; hidden: boolean; custom_height: boolean; style_id?: number }>
    columns: Array<{ column: number; end_column: number; width?: number; hidden: boolean; custom_width: boolean; best_fit: boolean; style_id?: number }>
    cells: Array<Record<string, unknown>>
    merged_ranges: Array<Record<string, unknown>>
  }>
  unsupported: NativeWorkbookUnsupportedV2[]
}

type MutableEffective = Record<string, unknown> & {
  number_format?: string; font_name?: string; font_size_points?: number; bold?: boolean; italic?: boolean
  font_color?: string; wrap_text?: boolean; shrink_to_fit?: boolean; text_rotation?: number
  horizontal_alignment?: string; vertical_alignment?: string
  projection: 'full' | 'partial'; unsupported: string[]
}

function fixture(): NativeWorkbookV2 {
  const workbook = JSON.parse(fixtureSource) as NativeWorkbookV2
  const mutable = workbook as unknown as MutableWorkbook
  mutable.normal_style = {
    style_xf_id: 0, font_id: 0, font_name: 'DejaVu Sans', font_size_points: 11, font_bold: false, font_italic: false,
    font_record_sha256: `sha256:${'e'.repeat(64)}`,
  }
  mutable.styles[0]!.effective = {
    ...mutable.styles[0]!.effective,
    font_name: 'DejaVu Sans',
    font_size_points: 11,
    bold: false,
    italic: false,
    font_color: '#000000',
    wrap_text: false,
    horizontal_alignment: 'general',
    vertical_alignment: 'bottom',
  }
  sealStyles(mutable)
  return workbook
}

function sealStyles(workbook: MutableWorkbook): void {
  for (const style of workbook.styles) style.raw_projection_sha256 = nativeWorkbookStyleRawProjectionSha256V2(style.effective as NativeWorkbookV2['styles'][number]['effective'])
}

function addStyle(workbook: MutableWorkbook, numberFormat: string): number {
  const id = workbook.styles.length
  const base = workbook.styles[0]!.effective
  workbook.styles.push({
    id,
    effective: {
      ...base,
      number_format: numberFormat,
    },
    raw_projection_sha256: '',
  })
  return id
}

function displayFixture(): MutableWorkbook {
  const workbook = fixture() as unknown as MutableWorkbook
  const isoDash = addStyle(workbook, 'yyyy-mm-dd')
  const isoSlash = addStyle(workbook, 'yyyy/mm/dd')
  const monthName = addStyle(workbook, 'mmmm d, yyyy')
  const systemDate = addStyle(workbook, '[$-F800]dddd, mmmm dd, yyyy')
  const integer = addStyle(workbook, '0')
  const fixed = addStyle(workbook, '0.00')
  workbook.sheets[0]!.columns.push({
    column: 6, end_column: 17, width: 24, hidden: false, custom_width: true, best_fit: false,
  })
  const cells = workbook.sheets[0]!.cells
  const push = (column: number, fields: Record<string, unknown>): void => {
    cells.push({ row: 0, column, ref: cellReference(column), style_id: 0, editable: true, ...fields })
  }
  push(6, { ooxml_type: 'b', value: { kind: 'boolean', storage: 'boolean', lexical: '1', rich: false } })
  push(7, { ooxml_type: 'b', value: { kind: 'boolean', storage: 'boolean', lexical: 'false', rich: false } })
  push(8, { ooxml_type: 'e', value: { kind: 'error', storage: 'error', lexical: '#DIV/0!', rich: false } })
  push(9, { ooxml_type: 'e', value: { kind: 'error', storage: 'error', lexical: '#N/A', rich: false } })
  push(10, { ooxml_type: 'd', style_id: isoDash, value: { kind: 'date', storage: 'date', lexical: '2026-08-27T14:03:04Z', rich: false } })
  push(11, { ooxml_type: 'd', style_id: isoSlash, value: { kind: 'date', storage: 'date', lexical: '2026-08-27T14:03:04Z', rich: false } })
  push(12, { ooxml_type: 'd', style_id: monthName, value: { kind: 'date', storage: 'date', lexical: '2026-08-27T14:03:04Z', rich: false } })
  push(13, { ooxml_type: 'd', style_id: systemDate, value: { kind: 'date', storage: 'date', lexical: '2026-08-27T14:03:04Z', rich: false } })
  push(14, { ooxml_type: 'n', formula: { text: '1+1', type: 'normal' } })
  push(15, { style_id: integer, value: { kind: 'number', storage: 'number', lexical: '001.2300', rich: false } })
  push(16, { style_id: fixed, value: { kind: 'number', storage: 'number', lexical: '001.2300', rich: false } })
  push(17, { style_id: monthName, value: { kind: 'number', storage: 'number', lexical: '44900', rich: false } })
  sealStyles(workbook)
  return workbook
}

function cellReference(column: number): string {
  let value = column + 1, name = ''
  while (value > 0) { value--; name = String.fromCharCode(65 + value % 26) + name; value = Math.floor(value / 26) }
  return `${name}1`
}

function metric(workbook: NativeWorkbookV2): NativeMaximumDigitWidthAuthorityV2 {
  return createNativeMaximumDigitWidthAuthorityV2(projectNativeWorkbookV2(workbook), FONT_BYTES)
}

function setup(workbook: NativeWorkbookV2 = fixture(), viewport = { row: 0, column: 0, end_row: 0, end_column: 5 }): { workbook: NativeWorkbookV2; model: NativeWorkbookRenderModelV2; geometry: NativeSheetGeometryV2 } {
  const model = projectNativeWorkbookV2(workbook)
  const geometry = compileNativeSheetGeometryV2(model, '7', viewport, metric(workbook))
  return { workbook, model, geometry }
}

function record(plan: NativeSheetCellPaintPlanV2): ReadonlyArray<NativeSheetCellPaintCommandV2> {
  const surface = createNativeSheetCellPaintRecordingSurfaceV2()
  emitNativeSheetCellPaintCommandsV2(plan, surface)
  return surface.finish()
}

function commandDigest(commands: ReadonlyArray<NativeSheetCellPaintCommandV2>): string {
  return createHash('sha256').update(JSON.stringify(commands)).digest('hex')
}

describe('native XLSX v2 cell glyph/display paint', () => {
  it('paints the lexical-render Hello string with integer-EMU glyph commands and a stable digest', () => {
    const { model, geometry } = setup()
    const before = JSON.stringify(model)
    const first = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    const second = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.paint_sha256).toBe('sha256:639db442a7c016dfdf07d64a521213aa555c21fc6b4ff1911224cad0a7533bfa')
    expect(first.capabilities).toEqual([{ name: 'native-cell-glyphs', level: 'exact' }])
    expect(first.gutter_emu).toBe(19_050)
    expect(first.coordinate_space).toBe('viewport-local')
    const hello = first.cells.find((cell) => cell.cell_ref === 'C1')
    expect(hello).toMatchObject({
      row: 0, column: 2, display_text: 'Hello', display_kind: 'string-literal',
      horizontal_alignment: 'left', vertical_alignment: 'bottom',
    })
    const glyphs = first.glyphs.filter((glyph) => glyph.cell_ref === 'C1')
    expect(hello?.glyph_end! - hello?.glyph_start!).toBe(5)
    expect(glyphs).toHaveLength(5)
    expect(glyphs.map((glyph) => glyph.glyph_id)).toEqual([43, 72, 79, 79, 82])
    expect(glyphs.every((glyph) => glyph.outline_kind === 'path' && glyph.path.length > 0 && glyph.fill_rgb === '#000000')).toBe(true)
    expect(glyphs.every((glyph) => Number.isSafeInteger(glyph.origin_x_emu) && Number.isSafeInteger(glyph.origin_y_emu))).toBe(true)
    expect(JSON.stringify(first)).not.toMatch(/css|px|canvas|dom/i)
    const commands = record(first)
    expect(commands[0]).toMatchObject({ kind: 'beginCellPaint', protocol: 'injoffice.xlsx.sheet-cell-paint', version: 1 })
    expect(commands.map((command) => command.kind)).toContain('fillGlyphPath')
    expect(commandDigest(commands)).toBe('4941a0a18e62215bebfa05d4088123274bd19085f59ef4df51bacdaa4b43b757')
    expect(commandDigest(commands)).toBe(commandDigest(record(second)))
    expect(Object.isFrozen(first)).toBe(true)
    expect(Object.isFrozen(first.glyphs[0]?.path)).toBe(true)
    expect(JSON.stringify(model)).toBe(before)
  })

  it('paints numeric lexical display and refuses General dates, rich, bold-mismatch, and uncached-formula cells atomically', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    const number = plan.cells.find((cell) => cell.cell_ref === 'A1')
    expect(number).toMatchObject({ display_text: '001.2300', display_kind: 'numeric-lexical', horizontal_alignment: 'right' })
    expect(plan.cells.some((cell) => cell.cell_ref === 'B1')).toBe(false)
    expect(plan.unsupported).toEqual(expect.arrayContaining([
      expect.objectContaining({ cell_ref: 'B1', code: 'CELL_FONT_MISMATCH' }),
      expect.objectContaining({ cell_ref: 'D1', code: 'CELL_DATE_DISPLAY' }),
      expect.objectContaining({ cell_ref: 'F1', code: 'CELL_RICH_TEXT' }),
    ]))
    expect(plan.glyphs.some((glyph) => glyph.cell_ref === 'F1' || glyph.cell_ref === 'D1')).toBe(false)
  })

  it('formats the locale-independent number/date subset without Date, Intl, or host TZ', () => {
    expect(formatNativeSheetCellDisplayV2('number', '001.2300', 'General')).toEqual({ status: 'ready', text: '001.2300' })
    expect(formatNativeSheetCellDisplayV2('number', '001.2300', '0')).toEqual({ status: 'ready', text: '1' })
    expect(formatNativeSheetCellDisplayV2('number', '001.2300', '0.00')).toEqual({ status: 'ready', text: '1.23' })
    expect(formatNativeSheetCellDisplayV2('number', '1.235', '0.00')).toEqual({ status: 'ready', text: '1.24' })
    expect(formatNativeSheetCellDisplayV2('number', '-0.5', '0')).toEqual({ status: 'ready', text: '-1' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', 'yyyy-mm-dd')).toEqual({ status: 'ready', text: '2026-08-27' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04+05:00', 'yyyy/mm/dd')).toEqual({ status: 'ready', text: '2026/08/27' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', 'm/d/yyyy')).toEqual({ status: 'ready', text: '8/27/2026' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', 'yyyy-mm-dd hh:mm:ss')).toEqual({ status: 'ready', text: '2026-08-27 14:03:04' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', '[$-409]mmmm d, yyyy')).toEqual({ status: 'ready', text: 'August 27, 2026' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', '[$-409]dddd')).toEqual({ status: 'ready', text: 'Thursday' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', 'General')).toMatchObject({ status: 'refused', code: 'CELL_DATE_DISPLAY' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', 'mmmm d, yyyy')).toMatchObject({ status: 'refused', code: 'CELL_DATE_DISPLAY' })
    expect(formatNativeSheetCellDisplayV2('date', '2026-08-27T14:03:04Z', '[$-F800]dddd, mmmm dd, yyyy')).toMatchObject({ status: 'refused', code: 'CELL_DATE_DISPLAY' })
    expect(formatNativeSheetCellDisplayV2('number', '44900', 'yyyy-mm-dd')).toMatchObject({ status: 'refused', code: 'CELL_DATE_DISPLAY' })
    expect(formatNativeSheetCellDisplayV2('number', '44900', 'yyyy-mm-dd', false)).toEqual({ status: 'ready', text: '2022-12-05' })
    expect(formatNativeSheetCellDisplayV2('number', '44900', 'm/d/yyyy', false)).toEqual({ status: 'ready', text: '12/5/2022' })
    expect(formatNativeSheetCellDisplayV2('number', '44900', 'yyyy/mm/dd', true)).toEqual({ status: 'ready', text: '2026/12/06' })
    expect(formatNativeSheetCellDisplayV2.toString()).not.toMatch(/\bIntl\b|\bnew Date\b/)
  })

  it('paints boolean TRUE/FALSE, error lexicals, and locale-independent dates as glyph paths', () => {
    const workbook = displayFixture()
    const { model, geometry } = setup(workbook as unknown as NativeWorkbookV2, { row: 0, column: 0, end_row: 0, end_column: 17 })
    const plan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    const truth = plan.cells.find((cell) => cell.cell_ref === 'G1')
    const falsy = plan.cells.find((cell) => cell.cell_ref === 'H1')
    const divZero = plan.cells.find((cell) => cell.cell_ref === 'I1')
    const na = plan.cells.find((cell) => cell.cell_ref === 'J1')
    const isoDash = plan.cells.find((cell) => cell.cell_ref === 'K1')
    const isoSlash = plan.cells.find((cell) => cell.cell_ref === 'L1')
    const integer = plan.cells.find((cell) => cell.cell_ref === 'P1')
    const fixed = plan.cells.find((cell) => cell.cell_ref === 'Q1')
    expect(truth).toMatchObject({ display_text: 'TRUE', display_kind: 'boolean-display', horizontal_alignment: 'left' })
    expect(falsy).toMatchObject({ display_text: 'FALSE', display_kind: 'boolean-display', horizontal_alignment: 'left' })
    expect(divZero).toMatchObject({ display_text: '#DIV/0!', display_kind: 'error-lexical', horizontal_alignment: 'left' })
    expect(na).toMatchObject({ display_text: '#N/A', display_kind: 'error-lexical', horizontal_alignment: 'left' })
    expect(isoDash).toMatchObject({ display_text: '2026-08-27', display_kind: 'date-display', horizontal_alignment: 'right' })
    expect(isoSlash).toMatchObject({ display_text: '2026/08/27', display_kind: 'date-display', horizontal_alignment: 'right' })
    expect(integer).toMatchObject({ display_text: '1', display_kind: 'numeric-lexical', horizontal_alignment: 'right' })
    expect(fixed).toMatchObject({ display_text: '1.23', display_kind: 'numeric-lexical', horizontal_alignment: 'right' })
    for (const cell of [truth, falsy, divZero, na, isoDash, isoSlash, integer, fixed]) {
      const glyphs = plan.glyphs.filter((glyph) => glyph.cell_ref === cell?.cell_ref)
      expect(glyphs.length).toBe(cell!.display_text.length)
      expect(glyphs.every((glyph) => glyph.outline_kind === 'path' && glyph.path.length > 0 && glyph.fill_rgb === '#000000')).toBe(true)
    }
    expect(plan.unsupported).toEqual(expect.arrayContaining([
      expect.objectContaining({ cell_ref: 'D1', code: 'CELL_DATE_DISPLAY' }),
      expect.objectContaining({ cell_ref: 'M1', code: 'CELL_DATE_DISPLAY' }),
      expect.objectContaining({ cell_ref: 'N1', code: 'CELL_DATE_DISPLAY' }),
      expect.objectContaining({ cell_ref: 'O1', code: 'CELL_FORMULA_NO_CACHE' }),
      expect.objectContaining({ cell_ref: 'R1', code: 'CELL_DATE_DISPLAY' }),
      expect.objectContaining({ cell_ref: 'F1', code: 'CELL_RICH_TEXT' }),
    ]))
    expect(plan.cells.some((cell) => cell.cell_ref === 'M1' || cell.cell_ref === 'O1' || cell.cell_ref === 'R1')).toBe(false)
    expect(plan.glyphs.some((glyph) => glyph.cell_ref === 'M1' || glyph.cell_ref === 'N1')).toBe(false)
  })

  it('paints BMP Unicode, combining clusters, and composite TrueType outlines the Normal font can shape', () => {
    const workbook = fixture() as unknown as MutableWorkbook
    workbook.sheets[0]!.columns.push({
      column: 6, end_column: 13, width: 24, hidden: false, custom_width: true, best_fit: false,
    })
    const cells = workbook.sheets[0]!.cells
    const push = (column: number, text: string): void => {
      cells.push({
        row: 0, column, ref: cellReference(column), style_id: 0, editable: true, ooxml_type: 'inlineStr',
        value: { kind: 'string', storage: 'inline', text, rich: false },
      })
    }
    push(6, 'Straße')
    push(7, '©€')
    push(8, 'e\u0301')
    push(9, 'e\u200D')
    push(10, 'e\uFE0F')
    push(11, '\uA7C7')
    push(12, 'e\u{1F600}')
    push(13, 'café')
    sealStyles(workbook)
    const { model, geometry } = setup(workbook as unknown as NativeWorkbookV2, { row: 0, column: 0, end_row: 0, end_column: 13 })
    const plan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    const street = plan.cells.find((cell) => cell.cell_ref === 'G1')
    const symbols = plan.cells.find((cell) => cell.cell_ref === 'H1')
    const combining = plan.cells.find((cell) => cell.cell_ref === 'I1')
    const cafe = plan.cells.find((cell) => cell.cell_ref === 'N1')
    expect(street).toMatchObject({ display_text: 'Straße', display_kind: 'string-literal', horizontal_alignment: 'left' })
    expect(symbols).toMatchObject({ display_text: '©€', display_kind: 'string-literal' })
    expect(combining).toMatchObject({ display_text: 'e\u0301', display_kind: 'string-literal' })
    expect(cafe).toMatchObject({ display_text: 'café', display_kind: 'string-literal' })
    expect(plan.glyphs.filter((glyph) => glyph.cell_ref === 'G1').map((glyph) => glyph.glyph_id)).toEqual([54, 87, 85, 68, 161, 72])
    expect(plan.glyphs.filter((glyph) => glyph.cell_ref === 'H1').map((glyph) => glyph.glyph_id)).toEqual([107, 2948])
    expect(plan.glyphs.filter((glyph) => glyph.cell_ref === 'I1').length).toBeGreaterThan(0)
    expect(plan.glyphs.filter((glyph) => glyph.cell_ref === 'N1').length).toBe(4)
    for (const cell of [street, symbols, combining, cafe]) {
      const glyphs = plan.glyphs.filter((glyph) => glyph.cell_ref === cell?.cell_ref)
      expect(glyphs.every((glyph) => glyph.outline_kind === 'path' && glyph.path.length > 0 && glyph.fill_rgb === '#000000')).toBe(true)
    }
    expect(plan.unsupported).toEqual(expect.arrayContaining([
      expect.objectContaining({ cell_ref: 'F1', code: 'CELL_RICH_TEXT' }),
      expect.objectContaining({ cell_ref: 'L1', code: 'CELL_MISSING_GLYPH' }),
    ]))
    expect(plan.unsupported.some((item) => item.cell_ref === 'I1' || item.cell_ref === 'N1')).toBe(false)
    expect(plan.cells.some((cell) => cell.cell_ref === 'F1' || cell.cell_ref === 'L1')).toBe(false)
    for (const ref of ['J1', 'K1', 'M1']) {
      const painted = plan.cells.find((cell) => cell.cell_ref === ref)
      const refused = plan.unsupported.find((item) => item.cell_ref === ref)
      expect(Boolean(painted) !== Boolean(refused)).toBe(true)
      if (refused) expect(['CELL_MISSING_GLYPH', 'CELL_COMPLEX_CLUSTER', 'CELL_COMPOSITE_OUTLINE']).toContain(refused.code)
      if (painted) expect(plan.glyphs.filter((glyph) => glyph.cell_ref === ref).every((glyph) => glyph.outline_kind === 'path' && glyph.path.length > 0)).toBe(true)
    }
  })

  it('validates the complete plan before the first push and validates replay before the first callback', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    const commands = record(plan)
    expect(commands[0]?.kind).toBe('beginCellPaint')
    expect(commands[1]?.kind).toBe('clipRect')
    expect(commands[commands.length - 1]?.kind).toBe('endCellPaint')
    const replayed: string[] = []
    replayNativeSheetCellPaintCommandsV2(replayed, commands, { execute(target, command) { target.push(command.kind) } })
    expect(replayed).toEqual(commands.map((command) => command.kind))

    const stalePlan = structuredClone(plan) as unknown as { glyphs: Array<{ fill_rgb: string }> }
    stalePlan.glyphs[0]!.fill_rgb = '#ABCDEF'
    let pushes = 0
    expect(() => emitNativeSheetCellPaintCommandsV2(stalePlan as unknown as NativeSheetCellPaintPlanV2, { push() { pushes++ } })).toThrowError(expect.objectContaining({ code: 'paint.planDigest' }))
    expect(pushes).toBe(0)

    const staleCommands = structuredClone(commands)
    const fill = staleCommands.find((command) => command.kind === 'fillGlyphPath') as Extract<NativeSheetCellPaintCommandV2, { kind: 'fillGlyphPath' }>
    ;(fill.glyph as { fill_rgb: string }).fill_rgb = '#ABCDEF'
    let callbacks = 0
    expect(() => replayNativeSheetCellPaintCommandsV2({}, staleCommands, { execute() { callbacks++ } })).toThrowError(expect.objectContaining({ code: 'paint.planDigest' }))
    expect(callbacks).toBe(0)

    const capacity = createNativeSheetCellPaintRecordingSurfaceV2(commands.length)
    emitNativeSheetCellPaintCommandsV2(plan, capacity, commands.length)
    expect(capacity.finish()).toHaveLength(commands.length)
    const tooSmall = createNativeSheetCellPaintRecordingSurfaceV2(commands.length - 1)
    expect(() => emitNativeSheetCellPaintCommandsV2(plan, tooSmall, commands.length)).toThrowError(expect.objectContaining({ code: 'paint.commandBudget' }))
    expect(tooSmall.commands).toEqual([])
  })

  it('paints wrap_text from HarfBuzz advances without rewriting display text', () => {
    const wrapText = 'Hello Hello Hello Hello Hello Hello'
    const wrapWorkbook = fixture() as unknown as MutableWorkbook
    const wrapStyle = addStyle(wrapWorkbook, 'General')
    wrapWorkbook.styles[wrapStyle]!.effective.wrap_text = true
    wrapWorkbook.sheets[0]!.rows.push({ row: 1, height_points: 72, hidden: false, custom_height: true })
    wrapWorkbook.sheets[0]!.cells.push({
      row: 1, column: 0, ref: 'A2', style_id: wrapStyle, editable: true, ooxml_type: 'inlineStr',
      value: { kind: 'string', storage: 'inline', text: wrapText, rich: false },
    })
    sealStyles(wrapWorkbook)
    const wrapNative = wrapWorkbook as unknown as NativeWorkbookV2
    const wrapSetup = setup(wrapNative, { row: 1, column: 0, end_row: 1, end_column: 0 })
    const wrapped = compileNativeSheetCellPaintV2(wrapSetup.model, wrapSetup.geometry, FONT_BYTES)
    const painted = wrapped.cells.find((cell) => cell.cell_ref === 'A2')
    const glyphs = wrapped.glyphs.filter((glyph) => glyph.cell_ref === 'A2')
    expect(painted).toMatchObject({ display_text: wrapText, display_kind: 'string-literal', horizontal_alignment: 'left' })
    expect(glyphs.length).toBeGreaterThan(5)
    expect(new Set(glyphs.map((glyph) => glyph.origin_y_emu)).size).toBeGreaterThan(1)
    expect(glyphs.every((glyph) => glyph.outline_kind === 'path' && glyph.path.length > 0)).toBe(true)
    expect(wrapped.unsupported.some((item) => item.cell_ref === 'A2')).toBe(false)

    const unwrappedWorkbook = fixture() as unknown as MutableWorkbook
    unwrappedWorkbook.sheets[0]!.rows.push({ row: 1, height_points: 72, hidden: false, custom_height: true })
    unwrappedWorkbook.sheets[0]!.cells.push({
      row: 1, column: 0, ref: 'A2', style_id: 0, editable: true, ooxml_type: 'inlineStr',
      value: { kind: 'string', storage: 'inline', text: wrapText, rich: false },
    })
    sealStyles(unwrappedWorkbook)
    const unwrappedNative = unwrappedWorkbook as unknown as NativeWorkbookV2
    const unwrappedSetup = setup(unwrappedNative, { row: 1, column: 0, end_row: 1, end_column: 0 })
    const unwrapped = compileNativeSheetCellPaintV2(unwrappedSetup.model, unwrappedSetup.geometry, FONT_BYTES)
    expect(unwrapped.cells.some((cell) => cell.cell_ref === 'A2')).toBe(false)
    expect(unwrapped.unsupported).toEqual(expect.arrayContaining([
      expect.objectContaining({ cell_ref: 'A2', code: 'CELL_OVERFLOW' }),
    ]))

    const shortRow = fixture() as unknown as MutableWorkbook
    const shortStyle = addStyle(shortRow, 'General')
    shortRow.styles[shortStyle]!.effective.wrap_text = true
    shortRow.sheets[0]!.rows.push({ row: 1, height_points: 15, hidden: false, custom_height: true })
    shortRow.sheets[0]!.cells.push({
      row: 1, column: 0, ref: 'A2', style_id: shortStyle, editable: true, ooxml_type: 'inlineStr',
      value: { kind: 'string', storage: 'inline', text: wrapText, rich: false },
    })
    sealStyles(shortRow)
    const shortNative = shortRow as unknown as NativeWorkbookV2
    const shortSetup = setup(shortNative, { row: 1, column: 0, end_row: 1, end_column: 0 })
    const shortPlan = compileNativeSheetCellPaintV2(shortSetup.model, shortSetup.geometry, FONT_BYTES)
    expect(shortPlan.cells.some((cell) => cell.cell_ref === 'A2')).toBe(false)
    expect(shortPlan.unsupported).toEqual(expect.arrayContaining([
      expect.objectContaining({ cell_ref: 'A2', code: 'CELL_OVERFLOW' }),
    ]))

    const singleLine = fixture() as unknown as MutableWorkbook
    singleLine.styles[0]!.effective.wrap_text = true
    sealStyles(singleLine)
    const { model, geometry } = setup(singleLine as unknown as NativeWorkbookV2)
    const singlePlan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    const hello = singlePlan.cells.find((cell) => cell.cell_ref === 'C1')
    expect(hello).toMatchObject({ display_text: 'Hello', display_kind: 'string-literal' })
    expect(singlePlan.glyphs.filter((glyph) => glyph.cell_ref === 'C1')).toHaveLength(5)
    expect(new Set(singlePlan.glyphs.filter((glyph) => glyph.cell_ref === 'C1').map((glyph) => glyph.origin_y_emu)).size).toBe(1)

    const wrapShrinkWorkbook = fixture() as unknown as MutableWorkbook
    const wrapShrinkStyle = addStyle(wrapShrinkWorkbook, 'General')
    wrapShrinkWorkbook.styles[wrapShrinkStyle]!.effective.wrap_text = true
    wrapShrinkWorkbook.styles[wrapShrinkStyle]!.effective.shrink_to_fit = true
    wrapShrinkWorkbook.sheets[0]!.rows.push({ row: 1, height_points: 72, hidden: false, custom_height: true })
    wrapShrinkWorkbook.sheets[0]!.cells.push({
      row: 1, column: 0, ref: 'A2', style_id: wrapShrinkStyle, editable: true, ooxml_type: 'inlineStr',
      value: { kind: 'string', storage: 'inline', text: wrapText, rich: false },
    })
    sealStyles(wrapShrinkWorkbook)
    const wrapShrinkSetup = setup(wrapShrinkWorkbook as unknown as NativeWorkbookV2, { row: 1, column: 0, end_row: 1, end_column: 0 })
    const wrapShrink = compileNativeSheetCellPaintV2(wrapShrinkSetup.model, wrapShrinkSetup.geometry, FONT_BYTES)
    const wrapShrinkCell = wrapShrink.cells.find((cell) => cell.cell_ref === 'A2')
    expect(wrapShrinkCell).toMatchObject({ display_text: wrapText, display_kind: 'string-literal' })
    expect(wrapShrink.unsupported.some((item) => item.cell_ref === 'A2')).toBe(false)
    expect(new Set(wrapShrink.glyphs.filter((glyph) => glyph.cell_ref === 'A2').map((glyph) => glyph.origin_y_emu)).size).toBeGreaterThan(1)
  })

  it('refuses theme fonts without exact bytes, merge non-origin text, and mismatched font bytes', () => {
    const excel = JSON.parse(excelFixtureSource) as NativeWorkbookV2
    expect(() => createNativeMaximumDigitWidthAuthorityV2(projectNativeWorkbookV2(excel), FONT_BYTES)).toThrowError(expect.objectContaining({ code: 'geometry.metricAuthority' }))

    const merged = fixture() as unknown as MutableWorkbook
    const c1 = merged.sheets[0]!.cells.find((cell) => cell.ref === 'C1')!
    const d1 = merged.sheets[0]!.cells.find((cell) => cell.ref === 'D1')!
    c1.editable = false
    delete d1.value
    delete d1.ooxml_type
    d1.editable = false
    merged.sheets[0]!.merged_ranges = [{ ref: 'C1:D1', row: 0, column: 2, end_row: 0, end_column: 3, editable: false }]
    const mergeLocation = ['MERGED_CELLS', 'merges', 'sheet:7', 'Worksheets/Sheet1.xml', '', ''].join('\0')
    merged.unsupported.push({
      id: `unsupported:${createHash('sha256').update(mergeLocation).digest('hex')}`,
      code: 'MERGED_CELLS', capability: 'merges', scope_id: 'sheet:7', part_name: 'Worksheets/Sheet1.xml',
      preservation: 'preserve-exact', message: 'merged source geometry',
    })
    const mergedNative = merged as unknown as NativeWorkbookV2
    const mergedModel = projectNativeWorkbookV2(mergedNative)
    const mergedGeometry = compileNativeSheetGeometryV2(mergedModel, '7', { row: 0, column: 0, end_row: 0, end_column: 5 }, metric(mergedNative))
    const mergedPlan = compileNativeSheetCellPaintV2(mergedModel, mergedGeometry, FONT_BYTES)
    const mergedHello = mergedPlan.cells.find((cell) => cell.cell_ref === 'C1')
    expect(mergedHello?.display_text).toBe('Hello')
    expect(mergedHello?.rect.width_emu).toBeGreaterThan(mergedGeometry.columns[2]!.width_emu)
    expect(mergedPlan.cells.some((cell) => cell.cell_ref === 'D1')).toBe(false)

    const { model, geometry } = setup()
    const tampered = FONT_BYTES.slice()
    tampered[100] = (tampered[100]! + 1) & 0xff
    expect(() => compileNativeSheetCellPaintV2(model, geometry, tampered)).toThrowError(expect.objectContaining({ code: 'paint.metricAuthority' }))
    expect(() => compileNativeSheetCellPaintV2(structuredClone(model), geometry, FONT_BYTES)).toThrowError(expect.objectContaining({ code: 'paint.modelAuthority' }))
  })

  it('refuses accessor and Proxy plans without invoking them', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    let getterCalls = 0
    const accessor = structuredClone(plan) as NativeSheetCellPaintPlanV2 & { trap?: string }
    Object.defineProperty(accessor, 'trap', { enumerable: true, get() { getterCalls++; return 'unsafe' } })
    expect(() => emitNativeSheetCellPaintCommandsV2(accessor, { push() {} })).toThrow(NativeSheetCellPaintError)
    expect(getterCalls).toBe(0)
    let reads = 0
    const proxied = new Proxy(plan, { get(target, property, receiver) { reads++; return Reflect.get(target, property, receiver) } })
    expect(() => emitNativeSheetCellPaintCommandsV2(proxied, { push() {} })).toThrow(NativeSheetCellPaintError)
    expect(reads).toBe(0)
    expect(() => validateNativeSheetCellPaintPlanV2({})).toThrowError(expect.objectContaining({ code: 'paint.planInvalid' }))
  })

  it('keeps Go extract fixtures valid and refuses invented workbook-level glyph-paint capabilities', () => {
    const extracted = JSON.parse(fixtureSource) as NativeWorkbookV2
    expect(extracted.capabilities.map((item) => item.name)).toEqual([
      'native-ooxml-parse', 'native-geometry', 'native-decorations', 'native-grid-commands', 'unsupported-content',
    ])
    expect(extracted.sheets[0]!.cells[2]).toMatchObject({ ref: 'C1', value: { text: 'Hello', rich: false } })
    const forged = structuredClone(extracted) as NativeWorkbookV2 & { capabilities: Array<{ name: string; level: string }> }
    forged.capabilities = [...forged.capabilities, { name: 'native-cell-glyphs', level: 'exact' }]
    expect(() => projectNativeWorkbookV2(forged)).toThrow(NativeWorkbookV2ValidationError)
  })

  it('paints exact rich runs, 1900/1904 serial dates, shrink-to-fit, and quadrant rotation', () => {
    const workbook = fixture() as unknown as MutableWorkbook
    workbook.date1904 = false
    const dateStyle = addStyle(workbook, 'yyyy-mm-dd')
    const shrinkStyle = addStyle(workbook, 'General')
    workbook.styles[shrinkStyle]!.effective.shrink_to_fit = true
    const rotateStyle = addStyle(workbook, 'General')
    workbook.styles[rotateStyle]!.effective.text_rotation = 90
    workbook.sheets[0]!.rows.push({ row: 1, height_points: 72, hidden: false, custom_height: true })
    workbook.sheets[0]!.columns.push({
      column: 6, end_column: 10, width: 24, hidden: false, custom_width: true, best_fit: false,
    })
    const cells = workbook.sheets[0]!.cells
    cells.push({
      row: 0, column: 6, ref: 'G1', style_id: 0, editable: false, ooxml_type: 'inlineStr',
      value: {
        kind: 'string', storage: 'inline', text: 'Hi', rich: true,
        runs: [
          { text: 'H', font_name: 'DejaVu Sans', bold: false, italic: false, font_size_points: 11, font_color: '#000000' },
          { text: 'i', font_name: 'DejaVu Sans', bold: false, italic: false, font_size_points: 11, font_color: '#112233' },
        ],
      },
    })
    const richLocation = ['RICH_CELL_STRING', 'rich-text', 'sheet:7', 'Worksheets/Sheet1.xml', 'G1', ''].join('\0')
    workbook.unsupported.push({
      id: `unsupported:${createHash('sha256').update(richLocation).digest('hex')}`,
      code: 'RICH_CELL_STRING', capability: 'rich-text', scope_id: 'sheet:7', part_name: 'Worksheets/Sheet1.xml',
      cell_ref: 'G1', preservation: 'preserve-exact', message: 'rich text runs remain source-authoritative',
    })
    cells.push({
      row: 0, column: 7, ref: 'H1', style_id: dateStyle, editable: true, ooxml_type: 'n',
      value: { kind: 'number', storage: 'number', lexical: '44900', rich: false },
    })
    cells.push({
      row: 1, column: 0, ref: 'A2', style_id: shrinkStyle, editable: true, ooxml_type: 'inlineStr',
      value: { kind: 'string', storage: 'inline', text: 'Hello Hello Hello Hello Hello Hello', rich: false },
    })
    cells.push({
      row: 1, column: 1, ref: 'B2', style_id: rotateStyle, editable: true, ooxml_type: 'inlineStr',
      value: { kind: 'string', storage: 'inline', text: 'Hi', rich: false },
    })
    sealStyles(workbook)
    const { model, geometry } = setup(workbook as unknown as NativeWorkbookV2, { row: 0, column: 0, end_row: 1, end_column: 10 })
    const plan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    const rich = plan.cells.find((cell) => cell.cell_ref === 'G1')
    const serial = plan.cells.find((cell) => cell.cell_ref === 'H1')
    const shrink = plan.cells.find((cell) => cell.cell_ref === 'A2')
    expect(rich).toMatchObject({ display_text: 'Hi', display_kind: 'string-literal' })
    expect(plan.glyphs.filter((glyph) => glyph.cell_ref === 'G1').map((glyph) => glyph.fill_rgb)).toEqual(['#000000', '#112233'])
    expect(serial).toMatchObject({ display_text: '2022-12-05', display_kind: 'date-display' })
    expect(shrink).toMatchObject({ display_text: 'Hello Hello Hello Hello Hello Hello' })
    expect(plan.unsupported.some((item) => item.cell_ref === 'G1' || item.cell_ref === 'H1' || item.cell_ref === 'A2')).toBe(false)
    const rotated = plan.cells.find((cell) => cell.cell_ref === 'B2')
    const rotatedRefusal = plan.unsupported.find((item) => item.cell_ref === 'B2')
    expect(Boolean(rotated) !== Boolean(rotatedRefusal)).toBe(true)
  })

  it('paints a second exact font face and still refuses substituting a missing theme face', () => {
    const boldBytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans-Bold.ttf')))
    const workbook = fixture() as unknown as MutableWorkbook
    workbook.styles[0]!.effective.font_name = 'DejaVu Sans'
    workbook.styles[0]!.effective.bold = true
    sealStyles(workbook)
    const { model, geometry } = setup(workbook as unknown as NativeWorkbookV2)
    const painted = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES, [boldBytes])
    const hello = painted.cells.find((cell) => cell.cell_ref === 'C1')
    expect(hello?.display_text).toBe('Hello')
    expect(painted.glyphs.filter((glyph) => glyph.cell_ref === 'C1').every((glyph) => glyph.face_id !== 'xlsx.normal-font')).toBe(true)

    const missing = fixture() as unknown as MutableWorkbook
    missing.styles[0]!.effective.font_name = 'ThemeDisplayFont'
    sealStyles(missing)
    const missingSetup = setup(missing as unknown as NativeWorkbookV2)
    const refused = compileNativeSheetCellPaintV2(missingSetup.model, missingSetup.geometry, FONT_BYTES, [boldBytes])
    expect(refused.unsupported).toEqual(expect.arrayContaining([
      expect.objectContaining({ cell_ref: 'C1', code: 'CELL_FONT_MISMATCH' }),
    ]))

    const keyed = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES, [{
      font_sha256: `sha256:${createHash('sha256').update(boldBytes).digest('hex')}`,
      bytes: boldBytes,
    }])
    expect(keyed.cells.find((cell) => cell.cell_ref === 'C1')?.display_text).toBe('Hello')
    expect(() => compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES, [{
      font_sha256: `sha256:${'a'.repeat(64)}`,
      bytes: boldBytes,
    }])).toThrowError(expect.objectContaining({ code: 'paint.metricAuthority' }))
  })

  it('paints reconstructible rich runs and numeric date tokens, and still refuses host-locale names and runless rich text', () => {
    const workbook = fixture() as unknown as MutableWorkbook
    const localeDate = addStyle(workbook, 'm/d/yyyy')
    const hostDate = addStyle(workbook, 'mmmm d, yyyy')
    workbook.date1904 = false
    workbook.sheets[0]!.columns.push({ column: 6, end_column: 9, width: 24, hidden: false, custom_width: true, best_fit: false })
    workbook.sheets[0]!.cells.push({
      row: 0, column: 6, ref: 'G1', style_id: 0, editable: false, ooxml_type: 'inlineStr',
      value: { kind: 'string', storage: 'inline', text: 'Hi', rich: true, runs: [{ text: 'Hi' }] },
    })
    const richLocation = ['RICH_CELL_STRING', 'rich-text', 'sheet:7', 'Worksheets/Sheet1.xml', 'G1', ''].join('\0')
    workbook.unsupported.push({
      id: `unsupported:${createHash('sha256').update(richLocation).digest('hex')}`,
      code: 'RICH_CELL_STRING', capability: 'rich-text', scope_id: 'sheet:7', part_name: 'Worksheets/Sheet1.xml',
      cell_ref: 'G1', preservation: 'preserve-exact', message: 'rich text runs remain source-authoritative',
    })
    workbook.sheets[0]!.cells.push({
      row: 0, column: 7, ref: 'H1', style_id: localeDate, editable: true, ooxml_type: 'd',
      value: { kind: 'date', storage: 'date', lexical: '2026-08-27T14:03:04Z', rich: false },
    })
    workbook.sheets[0]!.cells.push({
      row: 0, column: 8, ref: 'I1', style_id: hostDate, editable: true, ooxml_type: 'd',
      value: { kind: 'date', storage: 'date', lexical: '2026-08-27T14:03:04Z', rich: false },
    })
    sealStyles(workbook)
    const { model, geometry } = setup(workbook as unknown as NativeWorkbookV2, { row: 0, column: 0, end_row: 0, end_column: 9 })
    const plan = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    expect(plan.cells.find((cell) => cell.cell_ref === 'G1')).toMatchObject({ display_text: 'Hi', display_kind: 'string-literal' })
    expect(plan.cells.find((cell) => cell.cell_ref === 'H1')).toMatchObject({ display_text: '8/27/2026', display_kind: 'date-display' })
    expect(plan.glyphs.filter((glyph) => glyph.cell_ref === 'H1').length).toBe('8/27/2026'.length)
    expect(plan.unsupported).toEqual(expect.arrayContaining([
      expect.objectContaining({ cell_ref: 'F1', code: 'CELL_RICH_TEXT' }),
      expect.objectContaining({ cell_ref: 'I1', code: 'CELL_DATE_DISPLAY' }),
    ]))
    expect(plan.cells.some((cell) => cell.cell_ref === 'F1' || cell.cell_ref === 'I1')).toBe(false)
  })
})
