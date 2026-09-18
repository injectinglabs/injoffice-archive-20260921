import { describe, expect, it } from 'vitest'
import { placeNativeDocxFloatingTableV1 } from './nativeFloatingTableV1.js'
import type { NativeDocxTableFloatingPositionV1 } from './nativeContract.js'

const anchors = {
  page: { x_millipoints: 0, width_millipoints: 612_000 },
  margin: { x_millipoints: 72_000, width_millipoints: 468_000 },
  text: { x_millipoints: 306_000, width_millipoints: 234_000 },
}

function frame(overrides: Partial<NativeDocxTableFloatingPositionV1> = {}): NativeDocxTableFloatingPositionV1 {
  return { horizontal_anchor: 'margin', vertical_anchor: 'text', left_from_text_twips: 180, right_from_text_twips: 180, top_from_text_twips: 0, bottom_from_text_twips: 0, ...overrides }
}

describe('native DOCX floating table placement v1', () => {
  it('displaces the float from the top it would have had inline', () => {
    // floating-table-section-columns.docx: w:tblpY="-87" against the margin.
    expect(placeNativeDocxFloatingTableV1(frame({ y_twips: -87 }), anchors, 549_900)).toEqual({ x_millipoints: 72_000, y_offset_millipoints: -4_350 })
    // TC-table-DnD-move.docx: w:tblpY="69".
    expect(placeNativeDocxFloatingTableV1(frame({ y_twips: 69, left_from_text_twips: 141, right_from_text_twips: 141 }), anchors, 451_300)).toEqual({ x_millipoints: 72_000, y_offset_millipoints: 3_450 })
    // An absent w:tblpY leaves the float at its own inline top.
    expect(placeNativeDocxFloatingTableV1(frame(), anchors, 100_000).valueOf()).toEqual({ x_millipoints: 72_000, y_offset_millipoints: 0 })
  })

  it('takes the left edge from the anchor box the frame names', () => {
    // floatingtbl_with_formula.docx centres a 382.95 pt table on a 468 pt
    // margin box: Word's own export puts its border centres at 114.24..497.76.
    expect(placeNativeDocxFloatingTableV1(frame({ x_alignment: 'center', y_twips: 36 }), anchors, 382_950)).toEqual({ x_millipoints: 114_525, y_offset_millipoints: 1_800 })
    expect(placeNativeDocxFloatingTableV1(frame({ x_alignment: 'right' }), anchors, 100_000).valueOf()).toEqual({ x_millipoints: 440_000, y_offset_millipoints: 0 })
    expect(placeNativeDocxFloatingTableV1(frame({ horizontal_anchor: 'text', x_twips: 120 }), anchors, 100_000).valueOf()).toEqual({ x_millipoints: 312_000, y_offset_millipoints: 0 })
    expect(placeNativeDocxFloatingTableV1(frame({ horizontal_anchor: 'page', x_alignment: 'center' }), anchors, 100_000).valueOf()).toEqual({ x_millipoints: 256_000, y_offset_millipoints: 0 })
  })

  it('leaves a frame it cannot state exactly unplaced', () => {
    for (const unplaceable of [
      frame({ vertical_anchor: 'margin', y_twips: 100 }),
      frame({ vertical_anchor: 'page', y_twips: 100 }),
      frame({ y_alignment: 'center' }),
      frame({ top_from_text_twips: 120 }),
      frame({ bottom_from_text_twips: 120 }),
      frame({ x_alignment: 'inside' }),
      frame({ x_alignment: 'outside' }),
    ]) {
      expect(placeNativeDocxFloatingTableV1(unplaceable, anchors, 100_000)).toEqual({ unsupported: expect.any(String) })
    }
  })
})
