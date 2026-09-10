import type { NativeDocxLineFragmentV1 } from './nativeShapingLines.js'
import type { NativeDocxFillTextHighlightCommandV1 } from './nativePagePaintV1.js'

const COLORS: Readonly<Record<string, string>> = Object.freeze({ black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00', magenta: 'FF00FF', red: 'FF0000', yellow: 'FFFF00', white: 'FFFFFF', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0' })

/** Highlight geometry is derived from qualified run advances and font metrics,
 * never browser text measurement. This does not claim Word pixel equivalence. */
export function nativeTextHighlightCommandV1(highlight: string | undefined, fragment: NativeDocxLineFragmentV1, placedLineID: string, lineID: string, x: number, baselineY: number): { ok: true; command?: NativeDocxFillTextHighlightCommandV1 } | { ok: false; message: string } {
  if (!highlight || highlight === 'none') return { ok: true }
  if (!Object.hasOwn(COLORS, highlight)) return { ok: false, message: 'Highlight requires an explicit supported OOXML color' }
  if (fragment.source_kind !== 'run') return { ok: false, message: 'Highlight on list markers, controls and images is outside the text highlight subset' }
  const width = fragment.advance_inline_millipoints
  if (width === 0) return { ok: true }
  const y = baselineY - fragment.ascent_millipoints, height = fragment.ascent_millipoints - fragment.descent_millipoints
  if (![x,y,width,height,x+width,y+height].every((value) => Number.isSafeInteger(value) && value >= 0 && value <= 1_000_000_000_000) || width <= 0 || height <= 0) return { ok: false, message: 'Highlight geometry exceeds the bounded shaped-font metric rectangle' }
  return { ok: true, command: { kind: 'fill_text_highlight', id: `paint:${placedLineID}:${fragment.id}:highlight`, line_id: lineID, fragment_id: fragment.id, source_id: fragment.source_id, x_millipoints: x, y_millipoints: y, width_millipoints: width, height_millipoints: height, fill_rgb: COLORS[highlight]! } }
}
