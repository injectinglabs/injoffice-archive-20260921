import { formatNativeSheetCellDisplayV2,nativeTableNumberFormatPreview,formatNativeAccountingTextPreview,type NativeWorkbookObjectsV1 } from '@injoffice/sheets/browser'
import type { NativeCell, NativeWorkbook } from './nativeRoundTrip'

export type NativeCellPreview = { text: string; warning?: string; cached?: boolean }

/** Display only: never rewrite the value used by native mutations or evaluate formulas. */
export function nativeCellPreview(workbook: NativeWorkbook, cell: NativeCell | undefined,objects?:NativeWorkbookObjectsV1|null,sheetPart?:string): NativeCellPreview {
  if (!cell) return { text: '' }
  const value = cell.formula ? cell.formula.cached : cell.value
  if (!value) return cell.formula
    ? { text: '—', warning: 'No saved formula result. This preview does not recalculate formulas.' }
    : { text: '' }
  const cached = Boolean(cell.formula)
  const raw = value.text ?? value.lexical ?? ''
  if (value.kind === 'number' || value.kind === 'date') {
    const table=objects&&sheetPart?nativeTableNumberFormatPreview(objects,workbook.source.package_sha256,sheetPart,cell.row,cell.column,cell.style_id,value.kind,value.lexical??'',workbook.date1904):undefined
    if(table)return {text:table.text??raw,cached,...(table.warning?{warning:table.warning}:{})}
    const style = workbook.styles[cell.style_id]?.effective
    if (!style || style.number_format === undefined || style.unsupported.includes('number-format')) {
      return { text: raw, cached, warning: 'Number format unavailable; showing the stored value.' }
    }
    const display = formatNativeSheetCellDisplayV2(value.kind, value.lexical ?? '', style.number_format, workbook.date1904)
    if(display.status!=='ready'&&value.kind==='number'){
      const accounting=formatNativeAccountingTextPreview(value.lexical??'',style.number_format)
      if(accounting?.text!==undefined)return {text:accounting.text,cached,...(accounting.warning?{warning:accounting.warning}:{})}
    }
    return display.status === 'ready'
      ? { text: display.text, cached }
      : { text: raw, cached, warning: 'Unsupported number/date format; showing the stored value.' }
  }
  if (value.kind === 'boolean') return { text: value.lexical === '1' || value.lexical === 'true' ? 'TRUE' : 'FALSE', cached }
  return { text: raw, cached }
}
