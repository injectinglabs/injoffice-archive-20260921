import { describe, expect, it, vi } from 'vitest'
import { PrintManager, defaultPrintLayout, defaultPrintRender } from './manager'
import { hydrateNativePrintSetup, toNativePrintSetup } from './native'
import { validatePrintLayout, validatePrintRender } from './validation'

describe('print validation', () => {
  it('accepts defaults and rejects unsafe ranges, duplicate targets, and invalid custom values', () => {
    expect(validatePrintLayout(defaultPrintLayout('sheet-1')).ok).toBe(true)
    expect(validatePrintRender(defaultPrintRender()).ok).toBe(true)
    const invalid = defaultPrintLayout('sheet-1')
    invalid.subUnitIds = ['sheet-1', 'sheet-1']
    invalid.scale = 'Custom'
    invalid.customScale = 401
    invalid.margin = 'Custom'
    invalid.customMargins = { left: -1, right: 0, top: 0, bottom: 0, header: 0, footer: 0 }
    invalid.paperSize = 'Custom'
    invalid.pageSizeCustom = { width: 0, height: 10 }
    const result = validatePrintLayout(invalid)
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((entry) => entry.path)).toEqual(expect.arrayContaining(['/subUnitIds/1', '/customScale', '/customMargins/left', '/pageSizeCustom']))
  })

  it('returns diagnostics rather than throwing on malformed collection fields', () => {
    const layout = { ...defaultPrintLayout('sheet-1'), subUnitIds: null, freeze: null } as unknown as ReturnType<typeof defaultPrintLayout>
    const render = { ...defaultPrintRender(), headerFooter: null, headerFooterSetting: null } as unknown as ReturnType<typeof defaultPrintRender>
    expect(validatePrintLayout(layout)).toEqual(expect.objectContaining({ ok: false }))
    expect(validatePrintRender(render)).toEqual(expect.objectContaining({ ok: false }))
  })

  it('validates exact repeated-title ranges and typed preview watermarks', () => {
    const layout = defaultPrintLayout('sheet-1')
    layout.repeatRows = { startRow: 0, endRow: 1 }
    layout.repeatColumns = { startColumn: 0, endColumn: 2 }
    expect(validatePrintLayout(layout).ok).toBe(true)
    layout.repeatRows = { startRow: 4, endRow: 2 }
    expect(validatePrintLayout(layout)).toEqual(expect.objectContaining({ ok: false }))

    const render = defaultPrintRender()
    render.watermark = { kind: 'text', text: 'DRAFT', color: '#778899', opacity: 0.2, rotation: -45, fontSize: 42 }
    expect(validatePrintRender(render).ok).toBe(true)
    render.watermark = { kind: 'text', text: 'DRAFT', opacity: 2 }
    expect(validatePrintRender(render)).toEqual(expect.objectContaining({ ok: false }))
  })
})

describe('PrintManager', () => {
  it('runs cancellable dialog and print events in order', async () => {
    const print = vi.fn()
    const manager = new PrintManager({ print }, 'sheet-1')
    const events: string[] = []
    const stop = manager.onEvent((event) => events.push(event.name))
    expect(manager.openPrintDialog()).toBe(true)
    manager.updatePrintConfig({ direction: 'Landscape' })
    expect(await manager.print()).toBe(true)
    expect(print).toHaveBeenCalledWith(expect.objectContaining({ dialogOpen: true, layout: expect.objectContaining({ direction: 'Landscape' }) }))
    expect(events).toEqual(['before-open', 'opened', 'changed', 'before-confirm', 'confirmed'])
    stop()
  })

  it('allows before hooks to cancel without invoking the host', async () => {
    const print = vi.fn()
    const manager = new PrintManager({ print }, 'sheet-1')
    manager.onEvent((event) => { if (event.name === 'before-open' || event.name === 'before-confirm') event.cancel() })
    expect(manager.openPrintDialog()).toBe(false)
    expect(await manager.print()).toBe(false)
    expect(print).not.toHaveBeenCalled()
  })

  it('uses optional screenshot and clipboard adapters and rejects non-image output', async () => {
    const request = { subUnitId: 'sheet-1', range: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 } }
    const writeClipboardImage = vi.fn(() => true)
    const manager = new PrintManager({ print: vi.fn(), screenshot: () => 'data:image/png;base64,AA==', writeClipboardImage }, 'sheet-1')
    expect(await manager.saveScreenshotToClipboard(request)).toBe(true)
    expect(writeClipboardImage).toHaveBeenCalledWith('data:image/png;base64,AA==')
    const unavailable = new PrintManager({ print: vi.fn(), screenshot: () => 'https://example.invalid/image.png' }, 'sheet-1')
    expect(await unavailable.getScreenshot(request)).toBe(false)
  })

  it('restores validated durable settings atomically and isolates observers', () => {
    const manager = new PrintManager({ print: vi.fn() }, 'sheet-1')
    manager.openPrintDialog()
    manager.onEvent(() => { throw new Error('observer failure') })
    const replacement = manager.configurationSnapshot()
    replacement.layout.direction = 'Landscape'
    replacement.render.gridlines = true
    expect(manager.restoreConfiguration(replacement)).toBe(true)
    expect(manager.snapshot()).toMatchObject({ dialogOpen: true, layout: { direction: 'Landscape' }, render: { gridlines: true } })
    expect(manager.restoreConfiguration({ ...replacement, layout: { ...replacement.layout, paperSize: 'unknown' } })).toBe(false)
    expect(manager.configurationSnapshot().layout.paperSize).toBe('A4')
  })
})

describe('native print conversion', () => {
  it('maps the supported one-sheet subset to SetPrintSetup wire values', () => {
    const layout = defaultPrintLayout('sheet-1')
    layout.subUnitIds = [{ id: 'sheet-1', range: { startRow: 0, startColumn: 0, endRow: 19, endColumn: 5 } }]
    layout.direction = 'Landscape'
    layout.scale = 'FitPage'
    layout.fitToWidthPages = 2
    layout.fitToHeightPages = 3
    const render = defaultPrintRender()
    render.headerFooter = ['PageSize', 'WorksheetTitle']
    expect(toNativePrintSetup(layout, render, { sheetNameOf: () => 'Quarterly' })).toEqual({
      setup: expect.objectContaining({ sheetName: 'Quarterly', printAreaRef: 'A1:F20', orientation: 'landscape', fitToWidth: 2, fitToHeight: 3, headerCenter: 'Page &P of &N · &A' }),
      skipped: [],
    })
  })

  it('fails closed when requested options would be discarded', () => {
    const layout = defaultPrintLayout('sheet-1')
    layout.freeze = ['Row']
    const render = defaultPrintRender()
    render.gridlines = true
    render.hAlign = 'End'
    const result = toNativePrintSetup(layout, render, { sheetNameOf: () => 'Sheet1' })
    expect(result.setup).toBeNull()
    expect(result.skipped.join(' ')).toContain('repeated title')
    expect(result.skipped.join(' ')).toContain('end page alignment')
  })

  it('hydrates current native page setup, print area, and exact repeated-title ranges', () => {
    const result = hydrateNativePrintSetup({
      sheetName: 'Quarterly',
      printAreaRef: '$A$1:$F$20',
      repeatRowsRef: '$1:$2',
      repeatColumnsRef: '$A:$B',
      orientation: 'landscape',
      paperSize: 'Letter',
      fitToWidth: 2,
      fitToHeight: 0,
      margins: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
      oddHeader: '&CPage &P of &N · &A',
      horizontalCentered: true,
      verticalCentered: false,
      printGridlines: true,
      printHeadings: true,
    }, { sheetIdOf: (name) => name === 'Quarterly' ? 'sheet-1' : null })

    expect(result.value?.layout).toMatchObject({
      area: 'CurrentSelection', direction: 'Landscape', paperSize: 'Letter', scale: 'FitWidth', fitToWidthPages: 2, fitToHeightPages: 0, margin: 'Wide',
      freeze: ['Row', 'Column'], repeatRows: { startRow: 0, endRow: 1 }, repeatColumns: { startColumn: 0, endColumn: 1 },
      subUnitIds: [{ id: 'sheet-1', range: { startRow: 0, startColumn: 0, endRow: 19, endColumn: 5 } }],
    })
    expect(result.value?.render).toMatchObject({ gridlines: true, headings: true, hAlign: 'Middle', headerFooter: ['PageSize', 'WorksheetTitle'] })
    expect(result.warnings).toEqual([])
  })

  it('fails native hydration closed for unknown sheets and surfaces unsupported ranges', () => {
    expect(hydrateNativePrintSetup({
      sheetName: 'Missing', horizontalCentered: false, verticalCentered: false, printGridlines: false, printHeadings: false,
    }, { sheetIdOf: () => null }).value).toBeNull()
    const result = hydrateNativePrintSetup({
      sheetName: 'Data', printAreaRef: 'Data!A1:A2,Data!C1:C2', oddHeader: '&Lleft&Ccenter',
      horizontalCentered: false, verticalCentered: false, printGridlines: false, printHeadings: false,
    }, { sheetIdOf: () => 'sheet-1' })
    expect(result.value?.layout.area).toBe('CurrentSheet')
    expect(result.warnings.join(' ')).toContain('print area')
    expect(result.warnings.join(' ')).toContain('unsupported native section')
  })
})
