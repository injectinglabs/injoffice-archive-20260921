import type { NativeDocxTableFloatingPositionV1 } from './nativeContract.js'

/** The three boxes `w:horzAnchor` can name, in absolute page milli-points. */
export interface NativeDocxFloatingTableAnchorsV1 {
  page: { x_millipoints: number; width_millipoints: number }
  margin: { x_millipoints: number; width_millipoints: number }
  text: { x_millipoints: number; width_millipoints: number }
}

export interface NativeDocxFloatingTablePlacementV1 {
  /** Absolute page x of the floating table's left edge. */
  x_millipoints: number
  /** Signed shift from the top the table would have had inline at this position. */
  y_offset_millipoints: number
}

/**
 * Places a `w:tblpPr` frame this version can state exactly.
 *
 * `vertAnchor="text"` anchors the float to the top it would have had inline at
 * its own position in the flow -- after the preceding block and that block's
 * space-after -- displaced by `w:tblpY`. Measured against Microsoft Word
 * 16.112.4 PDF exports of the three hard-v2 packages that carry such a frame,
 * to 0.12 pt or better, the largest residual being one step of the exporter's
 * own 0.24 pt coordinate grid.
 *
 * Every other anchor, and any non-zero `w:topFromText`/`w:bottomFromText`, is
 * left unplaced: a keep-out distance above or below the float moves where the
 * following text resumes, and no reference measurement fixes that here.
 */
export function nativeDocxUnplaceableFloatingTableFrameV1(position: NativeDocxTableFloatingPositionV1): string | undefined {
  if (position.vertical_anchor !== 'text') return `vertical anchor ${position.vertical_anchor}`
  if (position.y_alignment !== undefined) return `vertical alignment ${position.y_alignment}`
  if (position.top_from_text_twips !== 0 || position.bottom_from_text_twips !== 0) return 'a non-zero keep-out distance above or below the frame'
  if (position.x_alignment !== undefined && !['left', 'center', 'right'].includes(position.x_alignment)) return `horizontal alignment ${position.x_alignment}`
  return undefined
}

export function placeNativeDocxFloatingTableV1(
  position: NativeDocxTableFloatingPositionV1,
  anchors: NativeDocxFloatingTableAnchorsV1,
  tableWidthMilliPoints: number,
): NativeDocxFloatingTablePlacementV1 | { unsupported: string } {
  const unplaceable = nativeDocxUnplaceableFloatingTableFrameV1(position)
  if (unplaceable !== undefined) return { unsupported: unplaceable }
  const anchor = anchors[position.horizontal_anchor]
  let x = anchor.x_millipoints + (position.x_twips ?? 0) * 50
  if (position.x_alignment === 'center') x = anchor.x_millipoints + Math.round((anchor.width_millipoints - tableWidthMilliPoints) / 2)
  else if (position.x_alignment === 'right') x = anchor.x_millipoints + anchor.width_millipoints - tableWidthMilliPoints
  if (!Number.isSafeInteger(x)) return { unsupported: 'a horizontal position outside the bounded coordinate range' }
  return { x_millipoints: x, y_offset_millipoints: (position.y_twips ?? 0) * 50 }
}
