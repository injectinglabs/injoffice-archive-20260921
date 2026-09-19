import type { XlsxNativeChart } from '@injoffice/xlsx-wasm';
import type { ChartAnchor, RangeRef } from '@injoffice/sheets/browser';
export const CHART_COLORS = ['#28764f', '#2459ad', '#b36c22', '#82569b', '#327e88', '#b34556', '#887225', '#526789'];
export function chartSelectionReason(range: RangeRef): string | undefined {
  if (range.end_row <= range.row || range.end_column <= range.column) return 'Select a header row, category column and at least one row of numeric values.';
  if (range.end_row - range.row > 1000 || range.end_column - range.column > 8) return 'Charts support up to 1,000 categories and 8 numeric series.';
}
export function chartAnchor(range: RangeRef): ChartAnchor {
  const from_column = Math.min(16375, range.end_column + 2), from_row = Math.min(1048557, range.row);
  return { from_row, from_column, to_row: from_row + 18, to_column: from_column + 8 };
}
/** Use a normalized domain to keep even finite 1e308/-1e308 source values finite. */
export function chartGeometry(chart: XlsxNativeChart) {
  if (!chart.editable || !chart.categories.length || !chart.series.length) throw new Error('Chart preview requires qualified literal data.');
  const values = chart.series.map(series => series.values.map(Number));
  let scale = 0;
  for (const series of values) for (const value of series) { if (!Number.isFinite(value)) throw new Error('Chart values must be finite.'); scale = Math.max(scale, Math.abs(value)); }
  if (!scale) scale = 1;
  const normalized = values.map(series => series.map(value => value / scale));
  let min = 0, max = 0;
  for (const series of normalized) for (const value of series) { min = Math.min(min, value); max = Math.max(max, value); }
  if (min === max) max = min + 1;
  const position = (value: number) => (value - min) / (max - min);
  const ticks = Array.from({length: 5}, (_, i) => { const value = min + (max - min) * i / 4; return { position: position(value), value: value * scale }; });
  return { normalized, position, zero: position(0), ticks };
}
export function chartNumber(value: number): string { return new Intl.NumberFormat('en', { maximumSignificantDigits: 4, notation: Math.abs(value) >= 1e9 || Math.abs(value) > 0 && Math.abs(value) < .001 ? 'scientific' : 'standard' }).format(value); }
