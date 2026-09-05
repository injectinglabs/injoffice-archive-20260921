import { OutlineManager } from '../../../packages/outlines/src/manager'
import { PrintConfigurationCommandController } from '../../../packages/print/src/commands'
import { PrintManager, defaultPrintLayout } from '../../../packages/print/src/manager'
import { PrintPreviewManager } from '../../../packages/print/src/preview'
import { extractSparklineValues } from '../../../packages/sparklines/src/extract'
import { compileSparklineGeometry, sparklineGeometryToSvg } from '../../../packages/sparklines/src/geometry'
import { SparklineManager } from '../../../packages/sparklines/src/manager'
import type { SparklineType } from '../../../packages/sparklines/src/types'
import { XlsxExchangeManager } from '../../../packages/xlsx-exchange/src/manager'
import type { XlsxExchangeJobSnapshot } from '../../../packages/xlsx-exchange/src/types'

export const SAMPLE_SPARKLINE_VALUES = [128, 156, 149, 188, 214, 246]

export function createPlaygroundSparkline(type: SparklineType, values: readonly (number | null)[]) {
  const manager = new SparklineManager({ idFactory: (kind) => `${kind}-demo` })
  const spec = manager.create({
    type,
    source: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 0, endColumn: Math.max(0, values.length - 1) },
    target: { sheetId: 'sheet-1', row: 1, column: 0 },
    options: { showHigh: true, showLow: true, showMarkers: type === 'line' },
  })
  const extracted = extractSparklineValues([values], spec.source, spec.options?.emptyCells)
  const geometry = compileSparklineGeometry({ type: spec.type, values: extracted, viewport: { width: 180, height: 48, padding: 3 }, options: spec.options })
  return { spec, extracted, svg: sparklineGeometryToSvg(geometry) }
}

export function createPlaygroundPrint(sheetId = 'sheet-1') {
  const printed: Array<{ paperSize: string; direction: string; area: string }> = []
  const manager = new PrintManager({
    print(snapshot) {
      printed.push({
        paperSize: snapshot.layout.paperSize,
        direction: snapshot.layout.direction,
        area: snapshot.layout.area,
      })
    },
  }, sheetId)
  return { manager, printed, layout: defaultPrintLayout(sheetId) }
}

export function createPlaygroundPrintWorkspace(sheetId = 'sheet-1') {
  const { manager, printed } = createPlaygroundPrint(sheetId)
  const controller = new PrintConfigurationCommandController(manager, { save() { /* session-only playground host */ } })
  const previewManager = new PrintPreviewManager<{ label: string }>({
    renderPreview: (request) => ({
      totalPages: 1,
      pages: [{
        number: 1,
        widthPoints: 612,
        heightPoints: 792,
        payload: { label: `${request.snapshot.layout.direction} ${request.snapshot.layout.paperSize}` },
      }],
    }),
  })
  return { manager, controller, previewManager, printed }
}

export function createPlaygroundOutlines() {
  const hidden: Array<{ axis: string; start: number; count: number; visible: boolean }> = []
  const manager = new OutlineManager({
    hide(_sheetId, axis, start, count) { hidden.push({ axis, start, count, visible: false }) },
    show(_sheetId, axis, start, count) { hidden.push({ axis, start, count, visible: true }) },
  })
  return { manager, hidden }
}

export function visibleOutlineRows(total: number, hiddenRanges: Array<{ start: number; count: number; visible: boolean }>): boolean[] {
  const visible = Array.from({ length: total }, () => true)
  for (const range of hiddenRanges) {
    for (let offset = 0; offset < range.count; offset++) {
      const index = range.start + offset
      if (index >= 0 && index < total) visible[index] = range.visible
    }
  }
  return visible
}

export function createPlaygroundExchange() {
  const jobs: XlsxExchangeJobSnapshot[] = []
  const manager = new XlsxExchangeManager<{ bytes: number; name: string }>({
    codec: {
      async importSnapshot(file, context) {
        context.report({ phase: 'converting', fraction: 0.6, totalBytes: file.bytes.byteLength })
        return { snapshot: { bytes: file.bytes.byteLength, name: file.name ?? 'workbook.xlsx' }, revision: `bytes:${file.bytes.byteLength}` }
      },
    },
  })
  manager.subscribe((event) => {
    if (event.type === 'job-created' || event.type === 'job-updated') jobs.push(event.job)
  })
  return { manager, jobs }
}
