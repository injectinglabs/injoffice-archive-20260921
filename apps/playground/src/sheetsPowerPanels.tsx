import { OutlineGutter } from '../../../packages/outlines/src/react'
import type { OutlineCommandController } from '../../../packages/outlines/src/commands'
import { SparklinePanel } from '../../../packages/sparklines/src/SparklinePanel'
import type { SparklineCommandController } from '../../../packages/sparklines/src/commands'
import { ShapePanel } from '@injoffice/shapes'
import type { ShapeManager } from '@injoffice/shapes'

export function SheetsPowerPanels({
  sparkline,
  outline,
  shapes,
  sheetId,
}: {
  sparkline: SparklineCommandController
  outline: OutlineCommandController
  shapes: ShapeManager
  sheetId: string
}) {
  return (
    <>
      <ShapePanel manager={shapes} />
      <SparklinePanel controller={sparkline} defaultSheetId={sheetId} />
      <OutlineGutter controller={outline} sheetId={sheetId} axis="row" />
      <OutlineGutter controller={outline} sheetId={sheetId} axis="column" />
    </>
  )
}
