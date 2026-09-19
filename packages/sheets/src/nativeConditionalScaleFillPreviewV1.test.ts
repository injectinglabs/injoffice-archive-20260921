import { describe, expect, it } from 'vitest'
import { decodeNativeConditionalScaleFillPreviewsV1, nativeConditionalScaleFillPreview } from './nativeConditionalScaleFillPreviewV1.js'
import type { NativeWorkbookObjectsV1 } from './nativeObjectsPreviewV1.js'

const REVISION = `sha256:${'a'.repeat(64)}`

function entry(overrides: Record<string, unknown> = {}) {
  return {
    sheet_id: '1',
    sheet_part: 'xl/worksheets/sheet1.xml',
    status: 'available',
    warnings: ['Read-only colour-scale preview.'],
    ranges: [{ ref: 'B3:B4', priority: 2, stops: 2 }],
    cells: [
      { row: 2, column: 1, color: '#FF0000' },
      { row: 3, column: 1, color: '#0000FF' },
    ],
    ...overrides,
  }
}

function objects(scales: unknown[]): NativeWorkbookObjectsV1 {
  return { protocol: 'injoffice.xlsx.preview-objects', version: 1, package_sha256: REVISION, tables: [], charts: [], conditional_scale_fills: scales } as unknown as NativeWorkbookObjectsV1
}

describe('colour-scale fill preview projection', () => {
  it('decodes an available overlay and an unavailable one', () => {
    const decoded = decodeNativeConditionalScaleFillPreviewsV1([entry(), { sheet_id: '2', sheet_part: 'xl/worksheets/sheet2.xml', status: 'unavailable', warnings: ['Colour-scale preview unavailable.'] }])
    expect(decoded).toHaveLength(2)
    expect(decoded[0]!.status).toBe('available')
    expect(decoded[0]!.cells).toHaveLength(2)
    expect(decoded[1]!.status).toBe('unavailable')
    expect(decoded[1]!.cells).toBeUndefined()
  })

  it('refuses a projection that is not exactly what the engine emits', () => {
    // Colours reach the raster unchanged, so a lowercase, short or
    // translucent value must never be accepted and painted as written.
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry({ cells: [{ row: 2, column: 1, color: '#ff0000' }] })])).toThrow()
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry({ cells: [{ row: 2, column: 1, color: 'red' }] })])).toThrow()
    // Out-of-order or duplicated cells would let one coordinate carry two
    // colours, and the index would silently keep the last one.
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry({ cells: [{ row: 3, column: 1, color: '#FF0000' }, { row: 2, column: 1, color: '#0000FF' }] })])).toThrow()
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry({ cells: [{ row: 2, column: 1, color: '#FF0000' }, { row: 2, column: 1, color: '#0000FF' }] })])).toThrow()
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry({ ranges: [{ ref: 'B3:B4', priority: 2, stops: 4 }] })])).toThrow()
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry({ warnings: [] })])).toThrow()
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry(), entry()])).toThrow()
    expect(() => decodeNativeConditionalScaleFillPreviewsV1([entry({ status: 'partial' })])).toThrow()
  })

  it('selects a colour by zero-based coordinate and only for the opened source', () => {
    const preview = objects([entry()])
    expect(nativeConditionalScaleFillPreview(preview, REVISION, 'xl/worksheets/sheet1.xml', 2, 1)).toBe('#FF0000')
    expect(nativeConditionalScaleFillPreview(preview, REVISION, 'xl/worksheets/sheet1.xml', 3, 1)).toBe('#0000FF')
    expect(nativeConditionalScaleFillPreview(preview, REVISION, 'xl/worksheets/sheet1.xml', 4, 1)).toBeUndefined()
    expect(nativeConditionalScaleFillPreview(preview, REVISION, 'xl/worksheets/sheet2.xml', 2, 1)).toBeUndefined()
    // A stale overlay must not paint over a different workbook.
    expect(nativeConditionalScaleFillPreview(preview, `sha256:${'b'.repeat(64)}`, 'xl/worksheets/sheet1.xml', 2, 1)).toBeUndefined()
  })

  it('paints nothing when the engine reported the sheet as unavailable', () => {
    const preview = objects([{ sheet_id: '1', sheet_part: 'xl/worksheets/sheet1.xml', status: 'unavailable', warnings: ['Colour-scale preview unavailable.'] }])
    expect(nativeConditionalScaleFillPreview(preview, REVISION, 'xl/worksheets/sheet1.xml', 2, 1)).toBeUndefined()
  })

  it('has no overlay when the engine emitted no colour scales', () => {
    expect(nativeConditionalScaleFillPreview(objects([]), REVISION, 'xl/worksheets/sheet1.xml', 0, 0)).toBeUndefined()
  })
})
