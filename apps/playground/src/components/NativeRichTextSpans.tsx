import type { NativeRichTextCellV1 } from '../../../../packages/sheets/src/nativeRichTextPreviewV1'
import { nativeRichTextRunDisclosureV1 } from '../../../../packages/sheets/src/nativeRichTextPreviewV1'

export interface NativeRichTextSpanBase {
  font_name?: string; font_size_points?: number; font_color?: string; bold?: boolean; italic?: boolean
}
/** Child spans of the existing clipped, positioned cell text. Host typography
 * stays approximate; each run independently resets to its cell's base font. */
export function NativeRichTextSpans({ entry, base, normal, loadedFont }: {
  entry: NativeRichTextCellV1; base: NativeRichTextSpanBase
  normal?: { font_name: string; font_bold?: boolean; font_italic?: boolean }; loadedFont?: string
}) {
  if (entry.status !== 'available' || !entry.runs) return <>{entry.text}</>
  return <>{entry.runs.map((run, i) => {
    const name = run.font_name ?? base.font_name ?? 'sans-serif'
    const bold = run.bold ?? base.bold ?? false, italic = run.italic ?? base.italic ?? false
    const family = loadedFont && name === normal?.font_name && bold === Boolean(normal.font_bold) && italic === Boolean(normal.font_italic) ? loadedFont : name
    return <tspan key={i} data-rich-run={i} data-rich-properties={run.properties} aria-label={nativeRichTextRunDisclosureV1(run)} fontFamily={family} fontSize={(run.font_size_points ?? base.font_size_points ?? 11) * 96 / 72} fontWeight={bold ? 700 : 400} fontStyle={italic ? 'italic' : 'normal'} fill={run.font_color ?? base.font_color ?? '#000000'} style={{ whiteSpace: 'pre' }}>{run.text}</tspan>
  })}</>
}
