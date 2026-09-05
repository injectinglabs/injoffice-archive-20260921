import type { Univer } from '@univerjs/core'
import { Subject } from 'rxjs'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_PRINT_CONFIGURATION_COMMANDS, INJOFFICE_PRINT_UI_COMMANDS, registerUniverPrintConfigurationCommands, registerUniverPrintUI } from './browser'
import { PrintManager } from './manager'

function harness() {
  const commands: Array<{ id: string; handler: (accessor?: unknown, params?: unknown) => unknown }> = []
  const undos: any[] = []
  const disposers: Array<ReturnType<typeof vi.fn>> = []
  const services = [
    { registerCommand: (command: any) => { commands.push(command); const dispose = vi.fn(); disposers.push(dispose); return { dispose } } },
    { pushUndoRedo: (record: any) => undos.push(record) },
  ]
  let service = 0
  return {
    univer: { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer,
    commands, undos, disposers,
  }
}

describe('Univer persisted print-configuration commands', () => {
  it('registers independent commands and pushes workbook-scoped atomic undo', async () => {
    const host = harness()
    const save = vi.fn()
    const manager = new PrintManager({ print: vi.fn() }, 'sheet-1')
    const registration = registerUniverPrintConfigurationCommands(host.univer, 'book-1', manager, { save }, { updateRender: false, replace: false })
    expect(host.commands.map(({ id }) => id).sort()).toEqual([
      INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore,
      INJOFFICE_PRINT_CONFIGURATION_COMMANDS.updateLayout,
    ].sort())

    const update = host.commands.find(({ id }) => id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.updateLayout)!
    await expect(update.handler(undefined, { patch: { direction: 'Landscape', margin: 'Narrow' } })).resolves.toBe(true)
    expect(manager.configurationSnapshot().layout).toMatchObject({ direction: 'Landscape', margin: 'Narrow' })
    expect(host.undos).toHaveLength(1)
    expect(host.undos[0]).toMatchObject({
      unitID: 'book-1',
      undoMutations: [{ id: INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore, params: { snapshot: { version: 1, layout: { direction: 'Portrait' } } } }],
      redoMutations: [{ id: INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore, params: { snapshot: { version: 1, layout: { direction: 'Landscape' } } } }],
    })
    const restore = host.commands.find(({ id }) => id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore)!
    await expect(restore.handler(undefined, host.undos[0].undoMutations[0].params)).resolves.toBe(true)
    expect(manager.configurationSnapshot().layout.direction).toBe('Portrait')
    expect(host.undos).toHaveLength(1)
    expect(save).toHaveBeenCalledTimes(2)

    registration.dispose()
    expect(host.disposers.every((dispose) => dispose.mock.calls.length === 1)).toBe(true)
  })

  it('rejects malformed parameters and converts persistence failure to command false', async () => {
    const host = harness()
    const manager = new PrintManager({ print: vi.fn() }, 'sheet-1')
    registerUniverPrintConfigurationCommands(host.univer, 'book-1', manager, { save: () => { throw new Error('offline') } })
    const layout = host.commands.find(({ id }) => id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.updateLayout)!
    const render = host.commands.find(({ id }) => id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.updateRender)!
    const replace = host.commands.find(({ id }) => id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.replace)!
    await expect(layout.handler(undefined, null)).resolves.toBe(false)
    await expect(render.handler(undefined, { patch: { gridlines: true } })).resolves.toBe(false)
    await expect(replace.handler(undefined, {})).resolves.toBe(false)
    expect(manager.configurationSnapshot().render.gridlines).toBe(false)
    expect(host.undos).toEqual([])
  })

  it('can route public commands and undo restores through a server-first target', async () => {
    const host = harness()
    const manager = new PrintManager({ print: vi.fn() }, 'sheet-1')
    const target = {
      updateLayout: vi.fn(async () => true),
      updateRender: vi.fn(async () => true),
      replace: vi.fn(async () => true),
      restore: vi.fn(async () => true),
      dispose: vi.fn(),
    }
    const factory = vi.fn(() => target)
    const registration = registerUniverPrintConfigurationCommands(host.univer, 'book-1', manager, { save: vi.fn() }, { commandTargetFactory: factory })
    const layout = host.commands.find(({ id }) => id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.updateLayout)!
    const restore = host.commands.find(({ id }) => id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore)!
    await expect(layout.handler(undefined, { patch: { paperSize: 'Letter' } })).resolves.toBe(true)
    await expect(restore.handler(undefined, { snapshot: manager.configurationSnapshot() })).resolves.toBe(true)
    expect(factory).toHaveBeenCalledOnce()
    expect(target.updateLayout).toHaveBeenCalledWith({ paperSize: 'Letter' })
    expect(target.restore).toHaveBeenCalledOnce()
    registration.dispose()
    expect(target.dispose).toHaveBeenCalledOnce()
  })
})

describe('Univer print UI entry', () => {
  it('registers a configurable public open command and Start ribbon item', async () => {
    const commands: any[] = []
    const menus: any[] = []
    const dispose = vi.fn()
    const services = [
      { registerCommand: (command: any) => { commands.push(command); return { dispose } } },
      { mergeMenu: (menu: any) => menus.push(menu) },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const open = vi.fn(() => true)
    const menuHidden$ = new Subject<boolean>()
    const registration = registerUniverPrintUI(univer, { open }, { title: 'Print workbook', tooltip: 'Check pages', order: 12, menuHidden$ })

    expect(commands.map(({ id }) => id)).toEqual([INJOFFICE_PRINT_UI_COMMANDS.open])
    await expect(commands[0].handler()).resolves.toBe(true)
    expect(open).toHaveBeenCalledOnce()
    expect(JSON.stringify(menus[0])).toContain(INJOFFICE_PRINT_UI_COMMANDS.open)
    const item = menus[0]['ribbon.start']['ribbon.start.layout'][INJOFFICE_PRINT_UI_COMMANDS.open].menuItemFactory()
    expect(item).toMatchObject({ title: 'Print workbook', tooltip: 'Check pages' })
    expect(item.hidden$).toBe(menuHidden$)
    registration.dispose()
    expect(dispose).toHaveBeenCalledOnce()
  })

  it('can omit both command and menu and contains host failures', async () => {
    const get = vi.fn()
    const univer = { __getInjector: () => ({ get }) } as unknown as Univer
    registerUniverPrintUI(univer, { open: vi.fn() }, { command: false })
    expect(get).not.toHaveBeenCalled()

    const commands: any[] = []
    const failing = { __getInjector: () => ({ get: () => ({ registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } }) }) } as unknown as Univer
    registerUniverPrintUI(failing, { open: () => { throw new Error('mount failed') } }, { menu: false })
    await expect(commands[0].handler()).resolves.toBe(false)
  })
})
