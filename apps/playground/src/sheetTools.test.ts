import { describe, expect, it } from 'vitest'
import {
  SAMPLE_SPARKLINE_VALUES,
  createPlaygroundExchange,
  createPlaygroundOutlines,
  createPlaygroundPrint,
  createPlaygroundPrintWorkspace,
  createPlaygroundSparkline,
  visibleOutlineRows,
} from './sheetTools'

describe('playground sheet tools', () => {
  it('compiles a sparkline spec to deterministic SVG', () => {
    const sparkline = createPlaygroundSparkline('line', SAMPLE_SPARKLINE_VALUES)
    expect(sparkline.spec.type).toBe('line')
    expect(sparkline.extracted).toEqual(SAMPLE_SPARKLINE_VALUES)
    expect(sparkline.svg).toContain('<svg')
    expect(sparkline.svg).toContain('<polyline')
  })

  it('runs print dialog events and records the host snapshot', async () => {
    const { manager, printed } = createPlaygroundPrint()
    expect(manager.openPrintDialog()).toBe(true)
    manager.updatePrintConfig({ direction: 'Landscape', paperSize: 'Letter' })
    expect(await manager.print()).toBe(true)
    expect(printed).toEqual([{ paperSize: 'Letter', direction: 'Landscape', area: 'CurrentSheet' }])
  })

  it('mounts a host-backed print workspace controller and preview', async () => {
    const { controller, previewManager, printed } = createPlaygroundPrintWorkspace()
    expect(await controller.updateLayout({ direction: 'Landscape' })).toBe(true)
    const session = previewManager.start('demo-print', controller.manager.snapshot())
    const document = await session.result
    expect(document.pages[0]?.payload.label).toContain('Landscape')
    expect(await controller.manager.print()).toBe(true)
    expect(printed[0]?.direction).toBe('Landscape')
  })

  it('collapses outline rows through the visibility adapter', () => {
    const { manager, hidden } = createPlaygroundOutlines()
    expect(manager.add({ id: 'q1', sheetId: 'sheet-1', axis: 'row', start: 2, end: 4, collapsed: true }).ok).toBe(true)
    expect(hidden.some((entry) => entry.start === 2 && entry.count === 3 && entry.visible === false)).toBe(true)
    const visible = visibleOutlineRows(6, hidden)
    expect(visible.slice(2, 5)).toEqual([false, false, false])
    expect(manager.setCollapsed('q1', false).ok).toBe(true)
  })

  it('imports bytes through the XLSX exchange job lifecycle', async () => {
    const { manager } = createPlaygroundExchange()
    const job = manager.importSnapshot({ source: { kind: 'bytes', bytes: new Uint8Array([1, 2, 3, 4]), name: 'demo.xlsx' } })
    const result = await job.result
    expect(result.snapshot).toEqual({ bytes: 4, name: 'demo.xlsx' })
    expect(job.snapshot().state).toBe('succeeded')
    expect(manager.listJobs()).toHaveLength(1)
  })
})
