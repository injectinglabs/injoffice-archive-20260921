import { compactNativeGeneralNumberPreviewV1, nativeTableNumberFormatPreview, type NativeWorkbookObjectsV1 } from '@injoffice/sheets/browser'
import { nativeCellPreview } from './nativeCellPreview'
import type { NativeCell, NativeWorkbook } from './nativeRoundTrip'

/** Host-only page display. The native model and mutation draft stay untouched. */
export function nativeSheetPageCellPreview(workbook: NativeWorkbook, cell: NativeCell | undefined, objects: NativeWorkbookObjectsV1, sheetPart: string, compactGeneral = false) {
  const display = nativeCellPreview(workbook, cell, objects, sheetPart)
  const value = cell?.formula ? cell.formula.cached : cell?.value
  const style = cell ? workbook.styles[cell.style_id]?.effective : undefined
  const warnings = display.warning ? [display.warning] : []
  // Match the native cell painter's General alignment policy using the stored
  // value kind, never by parsing formatted text or evaluating a formula.
  const alignmentUnavailable = style?.unsupported.some(flag => flag === 'horizontal-alignment' || flag === 'alignment-extended') ?? false
  const authoredAlignment = style?.horizontal_alignment ?? 'general'
  const horizontal = alignmentUnavailable || !style ? 'left' : authoredAlignment === 'general'
    ? value?.kind === 'number' || value?.kind === 'date' ? 'right' : 'left'
    : authoredAlignment
  if (alignmentUnavailable) warnings.push('Source alignment is unsupported; this preview uses left alignment without indent or rotation.')
  let text = display.text, compacted = false
  if (compactGeneral && cell && value?.kind === 'number' && style?.number_format === 'General' && !style.unsupported.includes('number-format')) {
    const table = nativeTableNumberFormatPreview(objects, workbook.source.package_sha256, sheetPart, cell.row, cell.column, cell.style_id, value.kind, value.lexical ?? '', workbook.date1904)
    if (!table) {
      const result = compactNativeGeneralNumberPreviewV1(value.lexical ?? '')
      if (result.status === 'ready') { text = result.text; compacted = text !== display.text }
      else warnings.push(result.warning)
    }
  }
  return { text, warnings, horizontal, cached: Boolean(cell?.formula && value), compacted, truncated: text.length > 2048, stored: value?.text ?? value?.lexical ?? '' }
}
