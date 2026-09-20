/** Excel's Number group over the native XLSX `style.patch` number_format, plus a fail-closed Borders control. */

/** Presets stay inside the qualified display subset (packages/sheets nativeCellDisplayV2) so the grid renders them rather than
 * falling back to "Stored value shown". The native transaction accepts any 1..255 character XML-safe format. */
export const NUMBER_FORMAT_PRESETS = { general: 'General', integer: '0', decimal: '0.00', currency: '"$"#,##0.00', percent: '0%', comma: '#,##0.00', date: 'yyyy-mm-dd' } as const;
export const MAX_DECIMAL_PLACES = 6;
export const DECIMALS_UNSUPPORTED = 'Decimal places apply to numeric formats with up to 6 places';
export const BORDERS_UNSUPPORTED = 'Borders are not supported by the native XLSX transaction';
const NUMERIC = /^(?:"([$£€¥] ?)")?(0|#,##0)(?:\.(0{1,6}))?(%)?(?:"( ?[$£€¥])")?$/;

export interface NumericFormat { prefix: string; grouped: boolean; decimals: number; percent: boolean; suffix: string }
export function parseNumericFormat(format: string | undefined): NumericFormat | undefined {
  if (!format || format === 'General') return { prefix: '', grouped: false, decimals: 0, percent: false, suffix: '' };
  const match = NUMERIC.exec(format); if (!match) return undefined;
  return { prefix: match[1] ?? '', grouped: match[2] === '#,##0', decimals: match[3]?.length ?? 0, percent: match[4] === '%', suffix: match[5] ?? '' };
}
export function numericFormatText(value: NumericFormat): string {
  return `${value.prefix ? `"${value.prefix}"` : ''}${value.grouped ? '#,##0' : '0'}${value.decimals ? `.${'0'.repeat(value.decimals)}` : ''}${value.percent ? '%' : ''}${value.suffix ? `"${value.suffix}"` : ''}`;
}
/** The next format with one more or one fewer decimal place, or undefined when the current format cannot be adjusted. */
export function adjustDecimals(format: string | undefined, delta: 1 | -1): string | undefined {
  const parsed = parseNumericFormat(format); if (!parsed) return undefined;
  const decimals = parsed.decimals + delta;
  if (decimals < 0 || decimals > MAX_DECIMAL_PLACES) return undefined;
  return numericFormatText({ ...parsed, decimals });
}
export function formatLabel(format: string): string {
  const labels: Record<string, string> = { General: 'General', '0': 'Integer', '0.00': 'Decimal · 2 places', '"$"#,##0.00': 'Currency', '0%': 'Percent', '#,##0.00': 'Comma · 2 places', 'yyyy-mm-dd': 'Date · YYYY-MM-DD' };
  return labels[format] ?? `Existing: ${format}`;
}

export default function SpreadsheetNumberFormat({ numberFormat, disabled, onChange }: { numberFormat?: string; disabled: boolean; onChange(numberFormat: string): void }) {
  const current = numberFormat || 'General', parsed = parseNumericFormat(current);
  const more = adjustDecimals(current, 1), fewer = adjustDecimals(current, -1);
  const options = [...new Set([...Object.values(NUMBER_FORMAT_PRESETS), current])];
  return <div className="sheet-number-format" aria-label="Number format">
    <select aria-label="Cell number format" disabled={disabled} value={current} onChange={event => onChange(event.target.value)}>{options.map(value => <option key={value} value={value}>{formatLabel(value)}</option>)}</select>
    <button aria-label="Currency format" title="Currency" aria-pressed={Boolean(parsed?.prefix)} disabled={disabled} onClick={() => onChange(NUMBER_FORMAT_PRESETS.currency)}>$</button>
    <button aria-label="Percent format" title="Percent" aria-pressed={Boolean(parsed?.percent)} disabled={disabled} onClick={() => onChange(NUMBER_FORMAT_PRESETS.percent)}>%</button>
    <button aria-label="Comma style" title="Comma style" aria-pressed={Boolean(parsed?.grouped && !parsed.prefix && !parsed.suffix)} disabled={disabled} onClick={() => onChange(NUMBER_FORMAT_PRESETS.comma)}>,</button>
    <button aria-label="Increase decimals" title={more ? 'Increase decimals' : DECIMALS_UNSUPPORTED} disabled={disabled || !more} onClick={() => more && onChange(more)}>+.0</button>
    <button aria-label="Decrease decimals" title={fewer ? 'Decrease decimals' : DECIMALS_UNSUPPORTED} disabled={disabled || !fewer} onClick={() => fewer && onChange(fewer)}>-.0</button>
  </div>;
}

/** Cell borders: the v1 StyleDelta has no border fields (borders are out of scope), so every option stays disabled with the reason. */
export function SpreadsheetBorders() {
  return <select className="sheet-borders" aria-label="Cell borders" title={BORDERS_UNSUPPORTED} disabled value="">
    <option value="">Borders</option>{['All borders', 'Outside borders', 'Top border', 'Bottom border', 'Left border', 'Right border', 'No border'].map(label => <option key={label} value={label}>{label}</option>)}
  </select>;
}
