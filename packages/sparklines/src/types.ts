export type SparklineType = 'line' | 'column' | 'win-loss'

export const SPARKLINE_TYPES: readonly SparklineType[] = ['line', 'column', 'win-loss']

/** Inclusive, zero-based range on one worksheet. Sparkline sources must be
 * one-dimensional: one row or one column. */
export interface SparklineRangeRef {
  sheetId: string
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

/** Zero-based cell containing the sparkline. */
export interface SparklineCellRef {
  sheetId: string
  row: number
  column: number
}

export type EmptyCellBehavior = 'gap' | 'zero' | 'connect'

export interface SparklineColors {
  series?: string
  negative?: string
  markers?: string
  high?: string
  low?: string
  first?: string
  last?: string
  axis?: string
}

export interface SparklineOptions {
  emptyCells?: EmptyCellBehavior
  rightToLeft?: boolean
  showMarkers?: boolean
  showHigh?: boolean
  showLow?: boolean
  showFirst?: boolean
  showLast?: boolean
  showNegative?: boolean
  lineWeight?: number
  /** Explicit bounds override individual or group-derived bounds. */
  min?: number
  max?: number
  colors?: SparklineColors
}

/** Plain JSON contract; no renderer or spreadsheet shell owns this model. */
export interface SparklineSpec {
  id: string
  type: SparklineType
  source: SparklineRangeRef
  target: SparklineCellRef
  groupId?: string
  options?: SparklineOptions
}

export interface SparklineGroup {
  id: string
  memberIds: string[]
}

export interface SparklineSnapshotV1 {
  version: 1
  sparklines: SparklineSpec[]
  groups: SparklineGroup[]
}

export interface SparklineViewport {
  width: number
  height: number
  padding?: number
}

export interface SparklinePoint {
  index: number
  value: number
  x: number
  y: number
  role: 'normal' | 'negative' | 'high' | 'low' | 'first' | 'last'
  color: string
}

export interface SparklinePath {
  points: Array<{ x: number; y: number }>
  color: string
  width: number
}

export interface SparklineBar {
  index: number
  value: number
  x: number
  y: number
  width: number
  height: number
  color: string
}

export interface SparklineGeometry {
  type: SparklineType
  width: number
  height: number
  domain: { min: number; max: number }
  baselineY: number
  paths: SparklinePath[]
  bars: SparklineBar[]
  markers: SparklinePoint[]
}

export type SparklineValueReader = (range: SparklineRangeRef) => readonly (readonly unknown[])[]
