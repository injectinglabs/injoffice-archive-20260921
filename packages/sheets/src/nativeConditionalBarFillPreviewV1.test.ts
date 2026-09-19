import { describe, expect, it } from 'vitest'
import { decodeNativeConditionalBarFillPreviewsV1, nativeConditionalBarFillPreview } from './nativeConditionalBarFillPreviewV1.js'
import type { NativeWorkbookObjectsV1 } from './nativeObjectsPreviewV1.js'

const REVISION = `sha256:${'a'.repeat(64)}`

function entry(overrides: Record<string, unknown> = {}) {
  return {
    sheet_id: '1',
    sheet_part: 'xl/worksheets/sheet1.xml',
    status: 'available',
    warnings: ['Read-only data-bar preview.'],
    ranges: [{ ref: 'A1:A2', priority: 2 }],
    cells: [
      { row: 0, column: 0, start_permille: 400, end_permille: 600, axis_permille: 400, color: '#0000FF', axis_color: '#000000' },
      { row: 1, column: 0, start_permille: 0, end_permille: 400, axis_permille: 400, color: '#FF0000', axis_color: '#000000' },
    ],
    ...overrides,
  }
}

function objects(bars: unknown[]): NativeWorkbookObjectsV1 {
  return { protocol: 'injoffice.xlsx.preview-objects', version: 1, package_sha256: REVISION, tables: [], charts: [], conditional_bar_fills: bars } as unknown as NativeWorkbookObjectsV1
}

describe('data-bar fill preview projection', () => {
  it('decodes an available overlay and an unavailable one', () => {
    const decoded = decodeNativeConditionalBarFillPreviewsV1([entry(), { sheet_id: '2', sheet_part: 'xl/worksheets/sheet2.xml', status: 'unavailable', warnings: ['Data-bar preview unavailable.'] }])
    expect(decoded).toHaveLength(2)
    expect(decoded[0]!.status).toBe('available')
    expect(decoded[0]!.cells![1]!.color).toBe('#FF0000')
    expect(decoded[1]!.cells).toBeUndefined()
  })

  it('accepts a bar with no axis and no border', () => {
    const decoded = decodeNativeConditionalBarFillPreviewsV1([entry({ cells: [{ row: 0, column: 0, start_permille: 0, end_permille: 250, axis_permille: -1, color: '#0000FF' }] })])
    expect(decoded[0]!.cells![0]!.axis_permille).toBe(-1)
    expect(decoded[0]!.cells![0]!.axis_color).toBeUndefined()
  })

  it('refuses a projection that is not exactly what the engine emits', () => {
    // A reversed span paints a negative-width rectangle, which renderers
    // disagree about; a span outside the cell paints over its neighbour.
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry({ cells: [{ row: 0, column: 0, start_permille: 600, end_permille: 400, axis_permille: -1, color: '#0000FF' }] })])).toThrow()
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry({ cells: [{ row: 0, column: 0, start_permille: 0, end_permille: 1200, axis_permille: -1, color: '#0000FF' }] })])).toThrow()
    // An axis position without its colour, or a colour without its position,
    // would draw a line in no colour or none at all.
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry({ cells: [{ row: 0, column: 0, start_permille: 0, end_permille: 250, axis_permille: 400, color: '#0000FF' }] })])).toThrow()
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry({ cells: [{ row: 0, column: 0, start_permille: 0, end_permille: 250, axis_permille: -1, color: '#0000FF', axis_color: '#000000' }] })])).toThrow()
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry({ cells: [{ row: 0, column: 0, start_permille: 0, end_permille: 250, axis_permille: -1, color: '#0000ff' }] })])).toThrow()
    // Out-of-order cells would let one coordinate carry two bars.
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry({ cells: [entry().cells[1], entry().cells[0]] })])).toThrow()
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry({ warnings: [] })])).toThrow()
    expect(() => decodeNativeConditionalBarFillPreviewsV1([entry(), entry()])).toThrow()
  })

  it('selects a bar by zero-based coordinate and only for the opened source', () => {
    const preview = objects([entry()])
    expect(nativeConditionalBarFillPreview(preview, REVISION, 'xl/worksheets/sheet1.xml', 0, 0)?.end_permille).toBe(600)
    expect(nativeConditionalBarFillPreview(preview, REVISION, 'xl/worksheets/sheet1.xml', 2, 0)).toBeUndefined()
    expect(nativeConditionalBarFillPreview(preview, REVISION, 'xl/worksheets/sheet2.xml', 0, 0)).toBeUndefined()
    expect(nativeConditionalBarFillPreview(preview, `sha256:${'b'.repeat(64)}`, 'xl/worksheets/sheet1.xml', 0, 0)).toBeUndefined()
  })

  it('paints nothing when the engine reported the sheet as unavailable', () => {
    const preview = objects([{ sheet_id: '1', sheet_part: 'xl/worksheets/sheet1.xml', status: 'unavailable', warnings: ['Data-bar preview unavailable.'] }])
    expect(nativeConditionalBarFillPreview(preview, REVISION, 'xl/worksheets/sheet1.xml', 0, 0)).toBeUndefined()
  })
})
