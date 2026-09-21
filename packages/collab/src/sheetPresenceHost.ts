import type { SheetPresencePopup } from './SheetPresenceLabel'

/** Disposable resources returned by the host's sheet facade. */
export interface SheetPresenceDisposable {
  dispose(): void
}

/** Zero-based inclusive coordinates supplied by the sheet host. */
export interface SheetPresenceCoordinates {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

interface SheetPresenceRange {
  getRange(): SheetPresenceCoordinates
  attachRangePopup(options: {
    componentKey: string
    direction: 'top-left'
    offset: [number, number]
    hideOnInvisible: boolean
    noPushMinimumGap: boolean
    zIndex: number
    extraProps: Record<string, unknown>
  }): SheetPresenceDisposable | null | undefined | void
}

interface SheetPresenceWorksheet {
  getSheetId(): string
  getRange(row: number, column: number, rows?: number, columns?: number): SheetPresenceRange
  getSelection(): { getActiveRange(): SheetPresenceRange | null | undefined } | null | undefined
  highlightRanges(ranges: SheetPresenceRange[], style: {
    stroke: string
    strokeWidth: number
    fill: string
    widgets: Record<string, never>
    autofillSize: number
  }, primary: null): SheetPresenceDisposable
}

interface SheetPresenceEvents {
  SelectionChanged: { worksheet?: { getSheetId(): string }; selections: SheetPresenceCoordinates[] }
  ActiveSheetChanged: object
  SheetEditChanging: {
    worksheet: { getSheetId(): string }
    row: number
    column: number
    value: { toPlainText(): string }
  }
  CommandExecuted: { id: string; params?: unknown }
}

/**
 * Sheet facade operations used by PresenceManager. A configured Univer facade
 * satisfies this structural contract; hosts own its installation and plugins.
 * Other collaboration features do not require a sheet editor dependency.
 */
export interface SheetPresenceHost {
  Event: { readonly [K in keyof SheetPresenceEvents]: K }
  addEvent<K extends keyof SheetPresenceEvents>(event: K, callback: (params: SheetPresenceEvents[K]) => void): SheetPresenceDisposable
  registerComponent(key: string, component: typeof SheetPresencePopup): SheetPresenceDisposable
  getActiveWorkbook(): {
    getId(): string
    getActiveSheet(): SheetPresenceWorksheet | null | undefined
  } | null | undefined
}
