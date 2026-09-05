import { SparklineCommandController } from '../../../packages/sparklines/src/commands'
import { SparklineManager } from '../../../packages/sparklines/src/manager'
import type { CreateSparklineInput } from '../../../packages/sparklines/src/manager'
import type { SparklineType } from '../../../packages/sparklines/src/types'
import { OutlineCommandController } from '../../../packages/outlines/src/commands'
import { OutlineManager } from '../../../packages/outlines/src/manager'
import type { OutlineAxis, OutlineGroup, OutlineVisibilityAdapter } from '../../../packages/outlines/src/types'

export interface SheetRange {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

export function sparklineInputFromRange(
  type: SparklineType,
  sheetId: string,
  range: SheetRange,
): CreateSparklineInput | null {
  if (range.endRow > range.startRow && range.endColumn > range.startColumn) return null
  return {
    type,
    source: {
      sheetId,
      startRow: range.startRow,
      startColumn: range.startColumn,
      endRow: range.endRow,
      endColumn: range.endColumn,
    },
    target: {
      sheetId,
      row: range.startRow,
      column: range.endColumn + 1,
    },
  }
}

export function outlineGroupFromRange(
  id: string,
  sheetId: string,
  axis: OutlineAxis,
  range: SheetRange,
): OutlineGroup {
  return {
    id,
    sheetId,
    axis,
    start: axis === 'row' ? range.startRow : range.startColumn,
    end: axis === 'row' ? range.endRow : range.endColumn,
    collapsed: false,
  }
}

export function createSheetsSparklineController(): SparklineCommandController {
  return new SparklineCommandController(new SparklineManager({ idFactory: (kind) => `${kind}-sheets` }))
}

export function createSheetsOutlineController(visibility?: OutlineVisibilityAdapter): OutlineCommandController {
  return new OutlineCommandController(new OutlineManager(visibility ?? { hide() {}, show() {} }))
}
