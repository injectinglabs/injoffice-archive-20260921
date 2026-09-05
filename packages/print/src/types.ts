export interface PrintRange {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

export type PrintArea = 'CurrentSheet' | 'Workbook' | 'CurrentSelection' | 'AllSelection'
export type PrintPaperSize = 'Letter' | 'Tabloid' | 'Legal' | 'Statement' | 'Executive' | 'Folio' | 'A3' | 'A4' | 'A5' | 'B4' | 'B5' | 'Custom'
export type PrintDirection = 'Portrait' | 'Landscape'
export type PrintScale = 'Origin' | 'FitWidth' | 'FitHeight' | 'FitPage' | 'Custom'
export type PrintFreeze = 'Row' | 'Column'
export type PrintPaperMargin = 'Normal' | 'Narrow' | 'Wide' | 'None' | 'Custom'
export type PrintAlign = 'Start' | 'End' | 'Middle'
export type PrintHeaderFooter = 'PageSize' | 'WorkbookTitle' | 'WorksheetTitle' | 'Date' | 'Time'

export interface PrintMargins {
  top: number
  right: number
  bottom: number
  left: number
  header: number
  footer: number
}

export interface PrintTarget {
  id: string
  range?: PrintRange
}

export interface PrintLayoutConfig {
  area: PrintArea
  subUnitIds: Array<string | PrintTarget>
  paperSize: PrintPaperSize
  direction: PrintDirection
  scale: PrintScale
  customScale: number
  /** Exact native fit width; used only by FitWidth/FitPage. */
  fitToWidthPages?: number
  /** Exact native fit height; used only by FitHeight/FitPage. Zero means unbounded. */
  fitToHeightPages?: number
  freeze: PrintFreeze[]
  /** Exact repeated worksheet-title rows when known. `freeze` remains the coarse UI vocabulary. */
  repeatRows?: { startRow: number; endRow: number }
  /** Exact repeated worksheet-title columns when known. `freeze` remains the coarse UI vocabulary. */
  repeatColumns?: { startColumn: number; endColumn: number }
  margin: PrintPaperMargin
  customMargins?: PrintMargins
  pageSizeCustom?: { width: number; height: number }
  maxRowsEachPage: number
  maxColumnsEachPage: number
}

export interface PrintHeaderFooterSetting {
  topLeft: string
  topCenter: string
  topRight: string
  bottomLeft: string
  bottomCenter: string
  bottomRight: string
}

export interface PrintRenderConfig {
  gridlines: boolean
  headings: boolean
  hAlign: PrintAlign
  vAlign: PrintAlign
  headerFooter: PrintHeaderFooter[]
  headerFooterSetting: PrintHeaderFooterSetting
  isCustomHeaderFooter?: boolean
  watermark?: PrintWatermark
}

export type PrintWatermark =
  | { kind: 'text'; text: string; color?: string; opacity?: number; rotation?: number; fontSize?: number }
  | { kind: 'image'; sourceId: string; opacity?: number; scale?: number }

export interface PrintSnapshot {
  layout: PrintLayoutConfig
  render: PrintRenderConfig
  dialogOpen: boolean
}

/** Durable print settings. Dialog visibility is deliberately not persisted or
 * placed on the undo stack. */
export interface PrintConfigurationSnapshotV1 {
  version: 1
  layout: PrintLayoutConfig
  render: PrintRenderConfig
}

export interface ScreenshotRequest {
  subUnitId: string
  range: PrintRange
  includeHeaders?: boolean
}

export interface PrintHost {
  print(snapshot: Readonly<PrintSnapshot>): Promise<void> | void
  screenshot?(request: Readonly<ScreenshotRequest>): Promise<string | false> | string | false
  writeClipboardImage?(dataUrl: string): Promise<boolean> | boolean
}

export type PrintEventName =
  | 'before-open'
  | 'opened'
  | 'before-confirm'
  | 'confirmed'
  | 'before-cancel'
  | 'canceled'
  | 'changed'

export interface PrintEvent {
  name: PrintEventName
  snapshot: Readonly<PrintSnapshot>
  cancelable: boolean
  canceled: boolean
  cancel(): void
}

export interface PrintIssue {
  path: string
  code: 'INVALID_ENUM' | 'INVALID_RANGE' | 'INVALID_NUMBER' | 'INVALID_TYPE' | 'REQUIRED' | 'DUPLICATE'
  message: string
}

export type PrintResult<T> = { ok: true; value: T } | { ok: false; issues: PrintIssue[] }
