import { formatNativeSheetCellDisplayV2 } from '@injoffice/sheets/browser'
import type { NativeCell, NativeWorkbook } from './nativeRoundTrip'

export type NativeCellPreview = { text: string; warning?: string; cached?: boolean }

/** Display only: never rewrite the value used by native mutations or evaluate formulas. */
export function nativeCellPreview(workbook: NativeWorkbook, cell: NativeCell | undefined): NativeCellPreview {
  if (!cell) return { text: '' }
  const value = cell.formula ? cell.formula.cached : cell.value
  if (!value) return cell.formula
    ? { text: '—', warning: 'No saved formula result. This preview does not recalculate formulas.' }
    : { text: '' }
  const cached = Boolean(cell.formula)
  const raw = value.text ?? value.lexical ?? ''
  if (value.kind === 'number' || value.kind === 'date') {
    const style = workbook.styles[cell.style_id]?.effective
    if (!style || style.number_format === undefined || style.unsupported.includes('number-format')) {
      return { text: raw, cached, warning: 'Number format unavailable; showing the stored value.' }
    }
    const display = formatNativeSheetCellDisplayV2(value.kind, value.lexical ?? '', style.number_format, workbook.date1904)
    return display.status === 'ready'
      ? { text: display.text, cached }
      : { text: raw, cached, warning: 'Unsupported number/date format; showing the stored value.' }
  }
  if (value.kind === 'boolean') return { text: value.lexical === '1' || value.lexical === 'true' ? 'TRUE' : 'FALSE', cached }
  return { text: raw, cached }
}
