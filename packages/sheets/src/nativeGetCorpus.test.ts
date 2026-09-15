import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import {
  compileNativeSheetCellPaintV2,
  compileNativeSheetGeometryV2,
  createNativeMaximumDigitWidthAuthorityV2,
  decodeNativeWorkbookV2,
  projectNativeWorkbookV2,
} from './index.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const corpusDir = resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-get-corpus')
const require = createRequire(import.meta.url)
const FONT_BYTES = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))

function loadWorkbook(id: string) {
  const decoded = decodeNativeWorkbookV2(readFileSync(resolve(corpusDir, `${id}.workbook.json`), 'utf8').trim())
  if (!decoded.ok) throw new Error(decoded.issues.map((issue) => `${issue.path}: ${issue.message}`).join('; '))
  return decoded.value
}

describe('native GET corpus cell-paint', () => {
  it('compiles the Excel-default + chart package that 503\'d Injecting GET', () => {
    const workbook = loadWorkbook('pass-excel-defaults-chart')
    expect(workbook.unsupported.some((item) => item.code === 'SHEET_VIEW_GEOMETRY')).toBe(false)
    expect(workbook.unsupported.some((item) => item.code === 'STYLE_RECORD_ATTRIBUTES')).toBe(false)
    expect(workbook.unsupported.some((item) => item.code === 'DRAWING_REFERENCE')).toBe(true)
    const covered = {
      ...workbook,
      unsupported: workbook.unsupported.filter((item) => item.code !== 'DRAWING_REFERENCE' && item.capability !== 'drawings'),
    }
    const model = projectNativeWorkbookV2(covered)
    const geometry = compileNativeSheetGeometryV2(
      model,
      covered.sheets[0]!.id,
      { row: 0, column: 0, end_row: 8, end_column: 10 },
      createNativeMaximumDigitWidthAuthorityV2(model, FONT_BYTES),
    )
    const paint = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    expect(paint.cells.length).toBeGreaterThan(0)
  })

  it('previews the frozen-pane package and records the pane as a typed read-only fact', () => {
    const workbook = loadWorkbook('refuse-freeze-pane')
    expect(workbook.unsupported.some((item) => item.code === 'SHEET_VIEW_GEOMETRY')).toBe(false)
    const pane = workbook.unsupported.find((item) => item.code === 'SHEET_VIEW_PANE')
    expect(pane?.scope_id).toBe(`sheet:${workbook.sheets[0]!.id}`)
    expect(pane?.message).toContain('1 frozen rows, 1 frozen columns')
    expect(workbook.sheets[0]!.sheet_view).toEqual({ pane_state: 'frozen', frozen_rows: 1, frozen_columns: 1, top_left_cell: 'B2', active_pane: 'bottomRight' })
    const model = projectNativeWorkbookV2(workbook)
    expect(model.sheets[0]!.sheet_view).toEqual(workbook.sheets[0]!.sheet_view)
    const geometry = compileNativeSheetGeometryV2(
      model,
      workbook.sheets[0]!.id,
      { row: 0, column: 0, end_row: 4, end_column: 2 },
      createNativeMaximumDigitWidthAuthorityV2(model, FONT_BYTES),
    )
    const paint = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    expect(paint.cells.length).toBeGreaterThan(0)
  })

  it('paints cellXf records that differ from their cellStyleXf parent without apply flags', () => {
    const workbook = loadWorkbook('refuse-apply-flags')
    const mismatch = workbook.unsupported.find((item) => item.code === 'STYLE_PARENT_APPLY_MISMATCH')
    expect(mismatch?.scope_id).toBe('workbook')
    expect(mismatch?.message).toContain('cellXf 1 fill id 1 differs from cellStyleXf id 0 without applyFill')
    expect(workbook.styles[1]!.effective.fill_color).toBe('#1F4E78')
    expect(workbook.styles[1]!.effective.bold).toBe(true)
    expect(workbook.styles[3]!.effective.number_format).toBe('$#,##0')
    const model = projectNativeWorkbookV2(workbook)
    const geometry = compileNativeSheetGeometryV2(
      model,
      workbook.sheets[0]!.id,
      { row: 0, column: 0, end_row: 4, end_column: 2 },
      createNativeMaximumDigitWidthAuthorityV2(model, FONT_BYTES),
    )
    const paint = compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)
    expect(paint.cells.length).toBeGreaterThan(0)
  })

  it('refuses showGridLines=0 as sheet view geometry', () => {
    const workbook = loadWorkbook('refuse-gridlines-off')
    const model = projectNativeWorkbookV2(workbook)
    expect(() => compileNativeSheetGeometryV2(
      model,
      workbook.sheets[0]!.id,
      { row: 0, column: 0, end_row: 4, end_column: 2 },
      createNativeMaximumDigitWidthAuthorityV2(model, FONT_BYTES),
    )).toThrow(/SHEET_VIEW_GEOMETRY/)
  })

  it('refuses quotePrefix=1 as style-record authority', () => {
    const workbook = loadWorkbook('refuse-quote-prefix')
    const model = projectNativeWorkbookV2(workbook)
    const geometry = compileNativeSheetGeometryV2(
      model,
      workbook.sheets[0]!.id,
      { row: 0, column: 0, end_row: 4, end_column: 2 },
      createNativeMaximumDigitWidthAuthorityV2(model, FONT_BYTES),
    )
    expect(() => compileNativeSheetCellPaintV2(model, geometry, FONT_BYTES)).toThrow(/STYLE_RECORD_ATTRIBUTES/)
  })
})
