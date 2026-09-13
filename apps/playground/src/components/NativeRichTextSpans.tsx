import type { NativeRichTextCellV1 } from '@injoffice/sheets/browser'
import { nativeRichTextRunDisclosureV1 } from '@injoffice/sheets/browser'

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
    return <tspan key={i} data-rich-run={i} data-rich-properties={run.properties} aria-label={nativeRichTextRunDisclosureV1(run)} fontFamily={family} fontSize={(run.font_size_points ?? base.font_size_points ?? 11) * 96 / 72} fontWeight={bold ? 700 : 400} fontStyle={italic ? 'italic' : 'normal'} fill={run.font_color ?? base.font_color ?? '#000000'} style={{ whiteSpace: 'pre', textDecoration: run.underline === 'single' ? 'underline solid' : 'none' }}>{run.text}</tspan>
  })}</>
}

/** Source samples use deliberate host boxes, independently of page geometry. */
export function NativeRichTextDetails({ entries, styles, warnings }: {
  entries: NativeRichTextCellV1[]; styles: readonly { id: number; effective: NativeRichTextSpanBase }[]; warnings: string[]
}) {
  return <section aria-label="Rich text run preview details">
    <h4>Source rich-text samples</h4>
    <p>These fixed browser boxes show stored text, direct run properties and qualified theme fonts. They are not worksheet positions or page geometry. Missing properties use each cell’s font as an approximate fallback; browser font matching and shaping may differ.</p>
    {warnings.map((warning, i) => <p key={i}>{warning}</p>)}
    {entries.map((entry, index) => <details key={`${entry.sheet_id}:${entry.ref}`} open={index < 3}>
      <summary>{entry.ref}: {entry.status === 'available' ? 'approximate runs' : 'run styling omitted'}</summary>
      <p>{entry.warnings.join(' ')}</p>
      <p>Source: {entry.sheet_part}; {entry.storage === 'shared' ? `${entry.source_part} shared string ${entry.shared_index}` : 'inline string'}; cell style {entry.style_id}.</p>
      {entry.status === 'available' && <svg role="img" aria-label={`Source rich-text sample ${entry.ref}`} viewBox="0 0 600 64" style={{ width: '100%', maxWidth: 600, overflow: 'hidden', border: '1px solid #bbb' }}>
        <text x={8} y={40}><NativeRichTextSpans entry={entry} base={styles.find(s => s.id === entry.style_id)?.effective ?? {}}/></text>
      </svg>}
      <p style={{ whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>Stored text: {entry.text}</p>
      {entry.runs && <ul>{entry.runs.map((run, i) => <li key={i}>Run {i + 1}: {nativeRichTextRunDisclosureV1(run)}</li>)}</ul>}
    </details>)}
  </section>
}
