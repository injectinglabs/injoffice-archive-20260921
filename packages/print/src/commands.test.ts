import { describe, expect, it, vi } from 'vitest'
import { PrintConfigurationCommandController, PrintConfigurationPersistenceError, type PrintConfigurationUndoRecord } from './commands'
import { PrintManager, defaultPrintLayout, defaultPrintRender } from './manager'
import type { PrintConfigurationSnapshotV1 } from './types'

function subject(save: (snapshot: Readonly<PrintConfigurationSnapshotV1>) => Promise<void> | void = () => {}) {
  const records: PrintConfigurationUndoRecord[] = []
  const manager = new PrintManager({ print: vi.fn() }, 'sheet-1')
  const controller = new PrintConfigurationCommandController(manager, { save }, { push: (record) => records.push(record) })
  return { manager, controller, records }
}

describe('PrintConfigurationCommandController', () => {
  it('persists the complete modeled layout and render surface before recording undo', async () => {
    const saved: Readonly<PrintConfigurationSnapshotV1>[] = []
    const { controller, records } = subject((snapshot) => { saved.push(snapshot) })
    expect(await controller.updateLayout({
      area: 'CurrentSelection',
      subUnitIds: [{ id: 'sheet-1', range: { startRow: 1, startColumn: 2, endRow: 20, endColumn: 8 } }],
      paperSize: 'Custom', pageSizeCustom: { width: 8.5, height: 13 }, direction: 'Landscape',
      scale: 'FitPage', customScale: 125, fitToWidthPages: 2, fitToHeightPages: 3,
      freeze: ['Row', 'Column'], repeatRows: { startRow: 0, endRow: 1 }, repeatColumns: { startColumn: 0, endColumn: 1 },
      margin: 'Custom', customMargins: { top: 1, right: 0.5, bottom: 1, left: 0.5, header: 0.25, footer: 0.25 },
      maxRowsEachPage: 50, maxColumnsEachPage: 10,
    })).toBe(true)
    expect(await controller.updateRender({
      gridlines: true, headings: true, hAlign: 'Middle', vAlign: 'End',
      headerFooter: ['PageSize', 'Date'], isCustomHeaderFooter: true,
      headerFooterSetting: { topLeft: 'Draft', topCenter: '&A', topRight: '', bottomLeft: '', bottomCenter: '&P', bottomRight: '' },
      watermark: { kind: 'text', text: 'INTERNAL', color: '#778899', opacity: 0.2, rotation: -45, fontSize: 36 },
    })).toBe(true)

    expect(saved).toHaveLength(2)
    expect(Object.isFrozen(saved[0].layout.subUnitIds)).toBe(true)
    expect(controller.snapshot()).toMatchObject({
      version: 1,
      layout: { area: 'CurrentSelection', paperSize: 'Custom', direction: 'Landscape', repeatRows: { startRow: 0, endRow: 1 } },
      render: { gridlines: true, headings: true, vAlign: 'End', watermark: { kind: 'text', text: 'INTERNAL' } },
    })
    expect(records.map(({ label }) => label)).toEqual(['Update print layout', 'Update print rendering'])
    expect(records[0].before.layout.direction).toBe('Portrait')
    expect(records[1].after.render.gridlines).toBe(true)
  })

  it('refuses malformed configuration without persistence, mutation, or undo', async () => {
    const save = vi.fn()
    const { controller, records } = subject(save)
    const before = controller.snapshot()
    expect(await controller.updateLayout(null as never)).toBe(false)
    expect(await controller.updateRender({ watermark: { kind: 'text', text: '' } } as never)).toBe(false)
    expect(await controller.replace({ version: 1, layout: { ...defaultPrintLayout('sheet-1'), direction: 'sideways' }, render: defaultPrintRender() })).toBe(false)
    expect(await controller.restore({ version: 2 })).toBe(false)
    expect(controller.snapshot()).toEqual(before)
    expect(save).not.toHaveBeenCalled()
    expect(records).toEqual([])
  })

  it('leaves local state and undo untouched when durable persistence rejects', async () => {
    const failure = new Error('store unavailable')
    const { controller, records } = subject(() => { throw failure })
    const before = controller.snapshot()
    const replacement = controller.snapshot()
    replacement.layout.direction = 'Landscape'
    replacement.render.gridlines = true
    await expect(controller.replace(replacement)).rejects.toMatchObject({
      name: 'PrintConfigurationPersistenceError',
      cause: failure,
    })
    expect(controller.snapshot()).toEqual(before)
    expect(records).toEqual([])
    expect(new PrintConfigurationPersistenceError(failure).cause).toBe(failure)
  })

  it('serializes concurrent saves so later patches compose from committed state', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const saved: PrintConfigurationSnapshotV1[] = []
    let call = 0
    const { controller } = subject(async (snapshot) => {
      if (call++ === 0) await gate
      saved.push(structuredClone(snapshot))
    })
    const first = controller.updateLayout({ direction: 'Landscape' })
    const second = controller.updateLayout({ paperSize: 'Letter' })
    await Promise.resolve()
    expect(saved).toEqual([])
    release()
    await expect(Promise.all([first, second])).resolves.toEqual([true, true])
    expect(saved[1].layout).toMatchObject({ direction: 'Landscape', paperSize: 'Letter' })
  })

  it('persists undo/redo restoration without recursively recording undo or changing dialog state', async () => {
    const save = vi.fn()
    const { manager, controller, records } = subject(save)
    manager.openPrintDialog()
    await controller.updateLayout({ direction: 'Landscape' })
    const before = records[0].before
    records.length = 0
    expect(await controller.restore(before)).toBe(true)
    expect(controller.snapshot().layout.direction).toBe('Portrait')
    expect(manager.snapshot().dialogOpen).toBe(true)
    expect(save).toHaveBeenCalledTimes(2)
    expect(records).toEqual([])
  })

  it('applies an authoritative snapshot through persistence with explicit undo policy', async () => {
    const save = vi.fn()
    const { controller, records } = subject(save)
    const replacement = controller.snapshot()
    replacement.layout.paperSize = 'Letter'
    expect(await controller.applyAuthoritative(replacement, { label: 'Collaborative print edit', recordUndo: true })).toBe(true)
    expect(save).toHaveBeenCalledOnce()
    expect(records).toHaveLength(1)
    expect(records[0]).toMatchObject({ label: 'Collaborative print edit', before: { layout: { paperSize: 'A4' } }, after: { layout: { paperSize: 'Letter' } } })

    const resync = controller.snapshot()
    resync.render.headings = true
    expect(await controller.applyAuthoritative(resync)).toBe(true)
    expect(records).toHaveLength(1)
  })
})
