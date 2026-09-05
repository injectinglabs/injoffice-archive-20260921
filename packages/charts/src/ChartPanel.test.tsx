import type { FUniver } from '@univerjs/core/lib/facade'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createChartPanelCommandBindings, downloadChartImageArtifact, exportChartFromPanel } from './ChartPanel'
import { ChartCommandController, type ChartUndoRecord } from './commands'
import { ChartImageExportManager } from './imageExport'
import { ChartManager } from './manager'
import type { ChartSpec } from './types'

function manager(): ChartManager {
  const sheet = {
    getSheetId: () => 'sheet-1',
    getRange: () => ({ getValues: () => [['Label', 'Value'], ['A', 1]] }),
    addFloatDomToPosition: (_config: unknown, id: string) => ({ id, dispose: () => undefined }),
    addFloatDomToRange: (_range: unknown, _config: unknown, _options: unknown, id: string) => ({ id, dispose: () => undefined }),
  }
  const workbook = {
    getActiveSheet: () => sheet,
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }
  return new ChartManager({ getActiveWorkbook: () => workbook } as unknown as FUniver)
}

function spec(id: string): ChartSpec {
  return {
    id,
    type: 'Column',
    title: id,
    range: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
  }
}

describe('ChartPanel integrations', () => {
  afterEach(() => vi.unstubAllGlobals())

  it('routes editing and layer controls through snapshot undo', () => {
    const records: ChartUndoRecord[] = []
    const controller = new ChartCommandController(manager(), { push: (record) => records.push(record) })
    controller.add(spec('chart-1')); controller.add(spec('chart-2')); records.length = 0
    const commands = createChartPanelCommandBindings(controller, 'chart-1')

    expect(commands.update({ title: 'Revenue' })).toBe(true)
    expect(commands.layer('bringToFront')).toBe(true)
    expect(commands.remove()).toBe(true)
    expect(records.map(({ label }) => label)).toEqual(['Update chart', 'Layer chart: bringToFront', 'Remove chart'])
    expect(records[1].after.charts.map(({ spec }) => spec.id)).toEqual(['chart-2', 'chart-1'])
  })

  it('exports a selected format and delivers the validated artifact', async () => {
    const chart = spec('chart-1')
    const render = vi.fn(async () => ({ mediaType: 'image/svg+xml' as const, bytes: new Uint8Array([60, 115, 118, 103, 62]), width: 640, height: 360 }))
    const exporter = new ChartImageExportManager((id) => id === chart.id ? chart : undefined, { render }, { createJobId: () => 'panel-export' })
    const deliver = vi.fn()

    await expect(exportChartFromPanel(exporter, chart.id, 'svg', deliver)).resolves.toMatchObject({
      chartId: chart.id,
      format: 'svg',
      mediaType: 'image/svg+xml',
    })
    expect(deliver).toHaveBeenCalledWith(expect.objectContaining({ jobId: 'panel-export', bytes: new Uint8Array([60, 115, 118, 103, 62]) }))
  })

  it('uses a sanitized filename and always revokes browser download URLs', () => {
    const anchor = { href: '', download: '', click: vi.fn(() => { throw new Error('blocked') }) }
    const createObjectURL = vi.fn(() => 'blob:chart')
    const revokeObjectURL = vi.fn()
    vi.stubGlobal('document', { createElement: () => anchor })
    vi.stubGlobal('URL', { createObjectURL, revokeObjectURL })

    expect(() => downloadChartImageArtifact({
      jobId: 'job-1', chartId: '../Quarterly Revenue', format: 'jpeg', mediaType: 'image/jpeg',
      bytes: new Uint8Array([1]), width: 1, height: 1,
    })).toThrow('blocked')
    expect(anchor.download).toBe('Quarterly-Revenue.jpg')
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:chart')
  })
})
