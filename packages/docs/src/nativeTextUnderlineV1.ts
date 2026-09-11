import type { NativeDocxLineFragmentV1 } from './nativeShapingLines.js'
import type { NativeDocxStrokeTextUnderlineCommandV1 } from './nativePagePaintV1.js'

/** Font-metric underline policy: post-table position/thickness; double rules
 * have one thickness of clear space. No browser or implicit font metrics. */
export function nativeTextUnderlineCommandsV1(style: string | undefined, color: string | undefined, fragment: NativeDocxLineFragmentV1, placement: string, line: string, x: number, baseline: number): { ok: true; commands: NativeDocxStrokeTextUnderlineCommandV1[] } | { ok: false; message: string } {
  if (!style || style === 'none') return { ok: true, commands: [] }
  if (!['single', 'double', 'words'].includes(style)) return { ok: false, message: 'Unsupported text underline style' }
  if (fragment.source_kind === 'image') return { ok: false, message: 'Image underline geometry is unsupported' }
  if (fragment.advance_inline_millipoints === 0 || style === 'words' && fragment.whitespace) return { ok: true, commands: [] }
  const position = fragment.underline_position_millipoints, thickness = fragment.underline_thickness_millipoints
  const fill = color ?? '000000'
  if (!fragment.face_id || !Number.isSafeInteger(position) || !Number.isSafeInteger(thickness) || thickness! <= 0 || !/^[0-9A-F]{6}$/.test(fill)) return { ok: false, message: 'Underline requires explicit content-addressed font position and thickness' }
  const result: NativeDocxStrokeTextUnderlineCommandV1[] = []
  for (let index = 0; index < (style === 'double' ? 2 : 1); index++) {
    const y = baseline - position! + index * 2 * thickness!, end = x + fragment.advance_inline_millipoints
    if (![x, end, y, thickness!, y - Math.ceil(thickness! / 2), y + Math.ceil(thickness! / 2)].every(value => Number.isSafeInteger(value) && value >= 0 && value <= 1_000_000_000_000) || end <= x) return { ok: false, message: 'Underline exceeds bounded font-metric coordinates' }
    result.push({ kind: 'stroke_text_underline', id: `paint:${placement}:${fragment.id}:underline:${index}`, line_id: line, fragment_id: fragment.id, source_id: fragment.source_id, stroke_index: index as 0 | 1, x1_millipoints: x, y1_millipoints: y, x2_millipoints: end, y2_millipoints: y, width_millipoints: thickness!, stroke_rgb: fill })
  }
  return { ok: true, commands: result }
}
