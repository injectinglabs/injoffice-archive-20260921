import type { PrintLayoutConfig, PrintMargins, PrintRange, PrintRenderConfig } from './types'
import { validatePrintLayout, validatePrintRender } from './validation'

export interface NativePrintSetup {
  sheetName: string
  printAreaRef: string
  orientation: 'portrait' | 'landscape'
  paperSize: Exclude<PrintLayoutConfig['paperSize'], 'Custom'>
  fitToWidth: number
  fitToHeight: number
  scale: number
  margins: { left: number; right: number; top: number; bottom: number; header: number; footer: number }
  headerCenter: string
  footerCenter: string
  horizontalCentered: boolean
  verticalCentered: boolean
  printGridlines: boolean
  printHeadings: boolean
}

export interface NativePrintContext {
  sheetNameOf(id: string): string | null
}

/** Bounded projection returned by xlsxpatch.ReadPrintSetups. Omitted values were absent in OOXML. */
export interface NativePrintState {
  sheetName: string
  printAreaRef?: string
  repeatRowsRef?: string
  repeatColumnsRef?: string
  orientation?: 'portrait' | 'landscape'
  paperSize?: Exclude<PrintLayoutConfig['paperSize'], 'Custom'>
  fitToWidth?: number
  fitToHeight?: number
  scale?: number
  margins?: NativePrintSetup['margins']
  oddHeader?: string
  oddFooter?: string
  horizontalCentered: boolean
  verticalCentered: boolean
  printGridlines: boolean
  printHeadings: boolean
  warnings?: string[]
}

export interface NativePrintHydrationContext {
  sheetIdOf(name: string): string | null
}

export interface NativePrintHydration {
  value: { layout: PrintLayoutConfig; render: PrintRenderConfig } | null
  warnings: string[]
}

const PRESET_MARGINS: Record<'Normal' | 'Narrow' | 'Wide' | 'None', PrintMargins> = {
  Normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  Narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  Wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  None: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
}

function columnName(index: number): string {
  let result = ''
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result
  return result
}

function a1(range: PrintRange): string {
  return `${columnName(range.startColumn)}${range.startRow + 1}:${columnName(range.endColumn)}${range.endRow + 1}`
}

function columnIndex(name: string): number {
  let value = 0
  for (const character of name.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value - 1
}

function parseA1(value: string): PrintRange | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(value)
  if (!match) return null
  const range = { startColumn: columnIndex(match[1]), startRow: Number(match[2]) - 1, endColumn: columnIndex(match[3]), endRow: Number(match[4]) - 1 }
  return range.startRow >= 0 && range.endRow >= range.startRow && range.endColumn >= range.startColumn && range.endRow <= 1_048_575 && range.endColumn <= 16_383 ? range : null
}

function parseRepeatRows(value: string | undefined): PrintLayoutConfig['repeatRows'] | null | undefined {
  if (value === undefined) return undefined
  const match = /^\$?(\d+):\$?(\d+)$/.exec(value)
  if (!match) return null
  const result = { startRow: Number(match[1]) - 1, endRow: Number(match[2]) - 1 }
  return result.startRow >= 0 && result.endRow >= result.startRow && result.endRow <= 1_048_575 ? result : null
}

function parseRepeatColumns(value: string | undefined): PrintLayoutConfig['repeatColumns'] | null | undefined {
  if (value === undefined) return undefined
  const match = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(value)
  if (!match) return null
  const result = { startColumn: columnIndex(match[1]), endColumn: columnIndex(match[2]) }
  return result.startColumn >= 0 && result.endColumn >= result.startColumn && result.endColumn <= 16_383 ? result : null
}

function tokens(render: PrintRenderConfig): { header: string; footer: string } {
  if (render.isCustomHeaderFooter) return { header: render.headerFooterSetting.topCenter, footer: render.headerFooterSetting.bottomCenter }
  const values = render.headerFooter.map((value) => ({ PageSize: 'Page &P of &N', WorkbookTitle: '&F', WorksheetTitle: '&A', Date: '&D', Time: '&T' })[value])
  return { header: values.join(' · '), footer: '' }
}

/** Convert only the subset SetPrintSetup can persist; unsupported choices are explicit diagnostics. */
export function toNativePrintSetup(layout: PrintLayoutConfig, render: PrintRenderConfig, context: NativePrintContext): { setup: NativePrintSetup | null; skipped: string[] } {
  const skipped: string[] = []
  if (layout.subUnitIds.length !== 1) skipped.push('native print setup currently accepts exactly one sheet target')
  const target = layout.subUnitIds[0]
  const id = typeof target === 'string' ? target : target?.id
  const sheetName = id ? context.sheetNameOf(id) : null
  if (!sheetName) skipped.push('print target sheet is unavailable')
  if (layout.area === 'Workbook' || layout.area === 'AllSelection') skipped.push(`native print setup cannot encode ${layout.area} as one worksheet setup`)
  if (layout.paperSize === 'Custom') skipped.push('native print setup does not yet persist custom paper dimensions')
  if (layout.freeze.length || layout.repeatRows || layout.repeatColumns) skipped.push('native print setup does not yet persist repeated title rows or columns')
  if (render.hAlign === 'End' || render.vAlign === 'End') skipped.push('native print setup cannot encode end page alignment')
  if (render.watermark !== undefined) skipped.push('watermarks are not persisted by the native print writer')
  if (render.isCustomHeaderFooter && [render.headerFooterSetting.topLeft, render.headerFooterSetting.topRight, render.headerFooterSetting.bottomLeft, render.headerFooterSetting.bottomRight].some(Boolean)) skipped.push('native print setup currently persists only centered headers and footers')
  if (skipped.length || !sheetName) return { setup: null, skipped }
  const range = typeof target === 'object' ? target.range : undefined
  const fitToWidth = layout.scale === 'FitWidth' || layout.scale === 'FitPage' ? layout.fitToWidthPages ?? 1 : 0
  const fitToHeight = layout.scale === 'FitHeight' || layout.scale === 'FitPage' ? layout.fitToHeightPages ?? 1 : 0
  const scale = layout.scale === 'Custom' ? layout.customScale : layout.scale === 'Origin' ? 100 : 0
  const content = tokens(render)
  const margins = layout.margin === 'Custom' ? layout.customMargins! : PRESET_MARGINS[layout.margin]
  return {
    setup: {
      sheetName,
      printAreaRef: range ? a1(range) : '',
      orientation: layout.direction === 'Landscape' ? 'landscape' : 'portrait',
      paperSize: layout.paperSize as Exclude<PrintLayoutConfig['paperSize'], 'Custom'>,
      fitToWidth,
      fitToHeight,
      scale,
      margins: { ...margins },
      headerCenter: content.header,
      footerCenter: content.footer,
      horizontalCentered: render.hAlign === 'Middle',
      verticalCentered: render.vAlign === 'Middle',
      printGridlines: render.gridlines,
      printHeadings: render.headings,
    },
    skipped,
  }
}

const PAPER_SIZES = new Set<PrintLayoutConfig['paperSize']>(['Letter', 'Tabloid', 'Legal', 'Statement', 'Executive', 'Folio', 'A3', 'A4', 'A5', 'B4', 'B5'])

function sameMargins(left: PrintMargins, right: PrintMargins): boolean {
  return (Object.keys(left) as Array<keyof PrintMargins>).every((key) => Math.abs(left[key] - right[key]) < 0.000_001)
}

function hydrateHeaderFooter(rawHeader: string | undefined, rawFooter: string | undefined, render: PrintRenderConfig, warnings: string[]): void {
  const centered = (value: string | undefined, label: string): string => {
    if (!value) return ''
    if (!value.startsWith('&C') || /&[LR]/.test(value.slice(2))) {
      warnings.push(`${label} uses unsupported native section or formatting syntax`)
      return ''
    }
    return value.slice(2)
  }
  const header = centered(rawHeader, 'odd header')
  const footer = centered(rawFooter, 'odd footer')
  const reverse = new Map<string, PrintRenderConfig['headerFooter'][number]>([
    ['Page &P of &N', 'PageSize'], ['&F', 'WorkbookTitle'], ['&A', 'WorksheetTitle'], ['&D', 'Date'], ['&T', 'Time'],
  ])
  const tokens = header ? header.split(' · ').map((value) => reverse.get(value)) : []
  if (header && !footer && tokens.every((value) => value !== undefined) && new Set(tokens).size === tokens.length) {
    render.headerFooter = tokens as PrintRenderConfig['headerFooter']
    return
  }
  if (header || footer) {
    render.isCustomHeaderFooter = true
    render.headerFooterSetting.topCenter = header
    render.headerFooterSetting.bottomCenter = footer
  }
}

/** Hydrate the representable current native setup into editable print configuration. */
export function hydrateNativePrintSetup(input: NativePrintState, context: NativePrintHydrationContext): NativePrintHydration {
  const warnings = Array.isArray(input?.warnings) ? input.warnings.filter((value): value is string => typeof value === 'string') : []
  if (!input || typeof input !== 'object' || typeof input.sheetName !== 'string') return { value: null, warnings: [...warnings, 'native print state is invalid'] }
  if (!context || typeof context.sheetIdOf !== 'function') return { value: null, warnings: [...warnings, 'native hydration context is invalid'] }
  const sheetId = context.sheetIdOf(input.sheetName)
  if (!sheetId) return { value: null, warnings: [...warnings, `native worksheet ${input.sheetName} is unavailable`] }
  const layout: PrintLayoutConfig = {
    area: 'CurrentSheet',
    subUnitIds: [sheetId],
    paperSize: 'A4', direction: 'Portrait', scale: 'FitWidth', customScale: 100,
    freeze: [], margin: 'Normal', maxRowsEachPage: 0, maxColumnsEachPage: 0,
  }
  for (const [name, value] of Object.entries({ horizontalCentered: input.horizontalCentered, verticalCentered: input.verticalCentered, printGridlines: input.printGridlines, printHeadings: input.printHeadings })) {
    if (typeof value !== 'boolean') warnings.push(`native ${name} flag is invalid and was ignored`)
  }
  const render: PrintRenderConfig = {
    gridlines: input.printGridlines === true, headings: input.printHeadings === true,
    hAlign: input.horizontalCentered === true ? 'Middle' : 'Start', vAlign: input.verticalCentered === true ? 'Middle' : 'Start',
    headerFooter: [], headerFooterSetting: { topLeft: '', topCenter: '', topRight: '', bottomLeft: '', bottomCenter: '', bottomRight: '' },
  }
  if (input.printAreaRef !== undefined) {
    const range = parseA1(input.printAreaRef)
    if (range) {
      layout.area = 'CurrentSelection'
      layout.subUnitIds = [{ id: sheetId, range }]
    } else warnings.push('native print area is not a representable single bounded range')
  }
  if (input.orientation !== undefined) {
    if (input.orientation === 'portrait' || input.orientation === 'landscape') layout.direction = input.orientation === 'landscape' ? 'Landscape' : 'Portrait'
    else warnings.push(`native orientation ${String(input.orientation)} is unsupported`)
  }
  if (input.paperSize !== undefined) {
    if (PAPER_SIZES.has(input.paperSize)) layout.paperSize = input.paperSize
    else warnings.push(`native paper size ${String(input.paperSize)} is unsupported`)
  }
  let width = input.fitToWidth ?? 0
  let height = input.fitToHeight ?? 0
  if (!Number.isSafeInteger(width) || width < 0 || !Number.isSafeInteger(height) || height < 0) {
    warnings.push('native fit dimensions are invalid and were ignored')
    width = height = 0
  }
  if (width > 0 || height > 0) {
    layout.scale = width > 0 && height > 0 ? 'FitPage' : width > 0 ? 'FitWidth' : 'FitHeight'
    if (input.fitToWidth !== undefined) layout.fitToWidthPages = width
    if (input.fitToHeight !== undefined) layout.fitToHeightPages = height
  } else if (input.scale !== undefined && Number.isSafeInteger(input.scale) && input.scale >= 10 && input.scale <= 400) {
    layout.scale = input.scale === 100 ? 'Origin' : 'Custom'
    layout.customScale = input.scale
  } else {
    if (input.scale !== undefined) warnings.push('native scale is invalid and was ignored')
    layout.scale = 'Origin'
  }
  if (input.margins) {
    const preset = (Object.entries(PRESET_MARGINS) as Array<[Exclude<PrintLayoutConfig['margin'], 'Custom'>, PrintMargins]>).find(([, value]) => sameMargins(value, input.margins!))
    if (preset) layout.margin = preset[0]
    else {
      layout.margin = 'Custom'
      layout.customMargins = { ...input.margins }
    }
  }
  const repeatRows = parseRepeatRows(input.repeatRowsRef)
  const repeatColumns = parseRepeatColumns(input.repeatColumnsRef)
  if (repeatRows === null) warnings.push('native repeated title rows are invalid or unsupported')
  else if (repeatRows) {
    layout.repeatRows = repeatRows
    layout.freeze.push('Row')
  }
  if (repeatColumns === null) warnings.push('native repeated title columns are invalid or unsupported')
  else if (repeatColumns) {
    layout.repeatColumns = repeatColumns
    layout.freeze.push('Column')
  }
  hydrateHeaderFooter(input.oddHeader, input.oddFooter, render, warnings)
  const checkedLayout = validatePrintLayout(layout)
  const checkedRender = validatePrintRender(render)
  if (!checkedLayout.ok || !checkedRender.ok) {
    const issues = [...checkedLayout.ok ? [] : checkedLayout.issues, ...checkedRender.ok ? [] : checkedRender.issues]
    return { value: null, warnings: [...warnings, ...issues.map((entry) => `hydrated ${entry.path}: ${entry.message}`)] }
  }
  return { value: { layout: checkedLayout.value, render: checkedRender.value }, warnings }
}
