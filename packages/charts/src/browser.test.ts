import type { Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { Subject } from 'rxjs'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_CHART_COMMANDS, registerUniverChartCommands } from './browser'

function apiHarness() {
  const sheet = {
    getSheetId: () => 'sheet-1',
    getActiveRange: () => ({ getRange: () => ({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 }) }),
    getRange: () => ({ getValues: () => [['Label', 'Value'], ['A', 1], ['B', 2]] }),
    addFloatDomToPosition: (_config: unknown, id: string) => ({ id, dispose: vi.fn() }),
  }
  return { getActiveWorkbook: () => ({
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getActiveRange: () => null,
    getSheetBySheetId: () => sheet,
  }) } as unknown as FUniver
}

describe('Univer chart commands', () => {
  it('registers configurable lifecycle commands, menu, and snapshot undo', () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const undos: any[] = []
    const menus: unknown[] = []
    const disposers: Array<ReturnType<typeof vi.fn>> = []
    const menuHidden$ = new Subject<boolean>()
    const services = [
      { registerCommand: (command: any) => { commands.push(command); const dispose = vi.fn(); disposers.push(dispose); return { dispose } } },
      { pushUndoRedo: (item: any) => undos.push(item) },
      { mergeMenu: (menu: unknown) => menus.push(menu) },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const registration = registerUniverChartCommands(univer, apiHarness(), {
      add: false,
      createInput: () => ({ type: 'Column', title: 'Revenue' }),
      layerHost: { setChartOrder: () => true },
      menuHidden$,
    })

    expect(commands.map(({ id }) => id).sort()).toEqual([
      INJOFFICE_CHART_COMMANDS.create,
      INJOFFICE_CHART_COMMANDS.layer,
      INJOFFICE_CHART_COMMANDS.remove,
      INJOFFICE_CHART_COMMANDS.restore,
      INJOFFICE_CHART_COMMANDS.update,
    ].sort())
    expect(menus).toHaveLength(1)
    const item = (menus[0] as any)['ribbon.insert']['ribbon.insert.media'][INJOFFICE_CHART_COMMANDS.create].menuItemFactory()
    expect(item.hidden$).toBe(menuHidden$)
    const create = commands.find(({ id }) => id === INJOFFICE_CHART_COMMANDS.create)!
    expect(create.handler()).toBe(true)
    expect(registration.controller.list()).toHaveLength(1)
    expect(undos).toHaveLength(1)
    expect(undos[0]).toMatchObject({
      unitID: 'book-1',
      undoMutations: [{ id: INJOFFICE_CHART_COMMANDS.restore, params: { snapshot: { version: 1, charts: [] } } }],
    })
    const restore = commands.find(({ id }) => id === INJOFFICE_CHART_COMMANDS.restore)!
    expect(restore.handler(undefined, undos[0].undoMutations[0].params)).toBe(true)
    expect(registration.controller.list()).toEqual([])
    registration.dispose()
    expect(disposers.every((dispose) => dispose.mock.calls.length === 1)).toBe(true)
  })

  it('omits Insert UI without a host input provider and rejects malformed command params', () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const services = [
      { registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: vi.fn() },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    expect(() => registerUniverChartCommands(univer, apiHarness())).not.toThrow()
    expect(commands.find(({ id }) => id === INJOFFICE_CHART_COMMANDS.create)?.handler(undefined, { type: 'NotAChart' })).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_CHART_COMMANDS.update)?.handler(undefined, {})).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_CHART_COMMANDS.remove)?.handler(undefined, {})).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_CHART_COMMANDS.layer)).toBeUndefined()
  })

  it('registers host-injected image export independently and can omit layering', async () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const services = [
      { registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: vi.fn() },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const render = vi.fn(async () => ({ mediaType: 'image/svg+xml' as const, bytes: new Uint8Array([60, 47, 62]), width: 320, height: 180 }))
    const events: string[] = []
    const registration = registerUniverChartCommands(univer, apiHarness(), {
      create: false,
      add: false,
      update: false,
      remove: false,
      layer: false,
      menu: false,
      imageExport: { host: { render }, onEvent: ({ type }) => events.push(type) },
    })
    registration.controller.add({
      id: 'chart-export',
      type: 'Line',
      range: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
    })

    expect(commands.map(({ id }) => id).sort()).toEqual([INJOFFICE_CHART_COMMANDS.exportImage, INJOFFICE_CHART_COMMANDS.restore].sort())
    const command = commands.find(({ id }) => id === INJOFFICE_CHART_COMMANDS.exportImage)!
    await expect(command.handler(undefined, { id: 'chart-export', options: { format: 'svg' } })).resolves.toMatchObject({ chartId: 'chart-export', format: 'svg', mediaType: 'image/svg+xml' })
    expect(events).toEqual(['started', 'completed'])
    expect(command.handler(undefined, {})).toBe(false)
    registration.dispose()
  })
})
