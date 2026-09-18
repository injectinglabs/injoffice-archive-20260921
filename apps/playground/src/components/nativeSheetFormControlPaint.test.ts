import { createRequire } from 'node:module'
import { readFileSync } from 'node:fs'
import { describe, it, expect } from 'vitest'
import { renderToStaticMarkup } from 'react-dom/server'
import { createElement } from 'react'
import {
  projectNativeWorkbookV2, createNativeMaximumDigitWidthAuthorityV2,
  compileNativeSheetGeometryV2, compileNativeSheetPagePreviewV1, layoutNativeFormControlsV1,
  type NativeWorkbookObjectsV1, type NativeFormControlV1,
} from '@injoffice/sheets/browser'
import { NativeSheetPageImages } from './NativeSheetPages'
import type { NativeWorkbook } from '../nativeRoundTrip'

const require = createRequire(import.meta.url)
const PART = 'Worksheets/Sheet1.xml'
/** Check Box 1 of the hard-v2 corpus file checkbox-form-control-align.xlsx. */
const CHECKBOX: NativeFormControlV1 = {
  sheet_id: '7', sheet_part: PART, ordinal: 1, shape_id: '1025', name: 'Check Box 1',
  kind: 'checkbox', object_type: 'CheckBox', checked: false,
  anchor: { kind: 'twoCellAnchor', from: { column: 0, row: 0, column_offset_emu: 123825, row_offset_emu: 57150 }, to: { column: 2, row: 1, column_offset_emu: 428625, row_offset_emu: 85725 } },
  caption: 'All effects', caption_size_points: 8, caption_align: 'right', caption_valign: 'bottom',
  warnings: ['Source anchor, control type and legacy caption only.'],
}

function paint(controls: NativeFormControlV1[]) {
  const workbook = JSON.parse(readFileSync(new URL('../../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json', import.meta.url), 'utf8')) as NativeWorkbook
  Object.assign(workbook, { normal_style: { style_xf_id: 0, font_id: 0, font_name: 'DejaVu Sans', font_size_points: 11, font_bold: false, font_italic: false, font_record_sha256: `sha256:${'e'.repeat(64)}` } })
  const model = projectNativeWorkbookV2(workbook)
  const font = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
  const authority = createNativeMaximumDigitWidthAuthorityV2(model, font)
  const geometry = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 2, end_column: 2 }, authority)
  const objects: NativeWorkbookObjectsV1 = {
    protocol: 'injoffice.xlsx.preview-objects', version: 1, package_sha256: model.source.package_sha256, tables: [], charts: [],
    page_settings: [{ sheet_id: '7', sheet_part: PART, status: 'available', warnings: ['Source settings'], settings: { paper: 'Letter', orientation: 'portrait', scale: 100, left_inches: 0.7, right_inches: 0.7, top_inches: 0.75, bottom_inches: 0.75 } }],
    form_controls: controls,
  }
  const plan = compileNativeSheetPagePreviewV1(geometry, objects)
  const formControls = layoutNativeFormControlsV1(geometry, objects)
  return renderToStaticMarkup(createElement(NativeSheetPageImages, {
    workbook, sheet: workbook.sheets[0]!, objects, geometry, plan, fontFamily: 'test-font', formControls,
  }))
}

describe('form control page paint', () => {
  it('strokes the checkbox square and its caption instead of leaving the page blank', () => {
    const markup = paint([structuredClone(CHECKBOX)])
    // The 12 pt glyph at 96 DPI is 16 CSS px, offset 15.1 pt and 7.5 pt into
    // the sheet, which is where Excel's own PDF paints it.
    expect(markup).toContain(`x="${15.1 * 12700 / 9525}" y="${7.5 * 12700 / 9525}" width="16" height="16" fill="#FFFFFF" stroke="#000000"`)
    expect(markup).toContain('All effects')
  })
  it('paints a check mark only for a checked control', () => {
    expect(paint([structuredClone(CHECKBOX)])).not.toContain('<path d="M ')
    expect(paint([{ ...structuredClone(CHECKBOX), checked: true }])).toContain('<path d="M ')
  })
  it('paints nothing at all for a control that is not a checkbox', () => {
    const markup = paint([{ ...structuredClone(CHECKBOX), kind: 'unsupported', object_type: 'Button', warnings: ['Form control "Button" is not a checkbox.'] }])
    expect(markup).not.toContain('All effects')
    expect(markup).not.toContain('stroke="#000000"')
  })
})
