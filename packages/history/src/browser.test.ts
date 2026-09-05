import type { Univer } from '@univerjs/core'
import { Subject } from 'rxjs'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_HISTORY_COMMANDS, INJOFFICE_HISTORY_UI_COMMANDS, registerUniverHistoryCommands, registerUniverHistoryUI } from './browser'
import type { HistoryCommandController } from './commands'

function harness() {
  const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
  const disposers: Array<ReturnType<typeof vi.fn>> = []
  const univer = {
    __getInjector: () => ({
      get: () => ({ registerCommand: (command: any) => {
        commands.push(command)
        const dispose = vi.fn()
        disposers.push(dispose)
        return { dispose }
      } }),
    }),
  } as unknown as Univer
  const controller = {
    list: vi.fn(async () => ({ versions: [] })),
    preview: vi.fn(async () => ({})),
    capture: vi.fn(async () => ({})),
    restore: vi.fn(async () => ({})),
    cancel: vi.fn(() => true),
  } as unknown as HistoryCommandController<unknown>
  return { commands, controller, disposers, univer }
}

describe('Univer history commands', () => {
  it('registers configurable workflows and disposes them', async () => {
    const { commands, controller, disposers, univer } = harness()
    const registration = registerUniverHistoryCommands(univer, controller, { list: false })
    expect(commands.map(({ id }) => id).sort()).toEqual([
      INJOFFICE_HISTORY_COMMANDS.preview,
      INJOFFICE_HISTORY_COMMANDS.capture,
      INJOFFICE_HISTORY_COMMANDS.restore,
      INJOFFICE_HISTORY_COMMANDS.cancel,
    ].sort())

    const byId = (id: string) => commands.find((entry) => entry.id === id)!
    await expect(byId(INJOFFICE_HISTORY_COMMANDS.preview).handler(undefined, { versionId: 'v1' })).resolves.toBe(true)
    await expect(byId(INJOFFICE_HISTORY_COMMANDS.capture).handler(undefined, { author: { id: 'agent-1', kind: 'agent' } })).resolves.toBe(true)
    await expect(byId(INJOFFICE_HISTORY_COMMANDS.restore).handler(undefined, { sourceVersionId: 'v1', author: { id: 'u1', kind: 'user' } })).resolves.toBe(true)
    expect(byId(INJOFFICE_HISTORY_COMMANDS.cancel).handler()).toBe(true)
    expect(controller.preview).toHaveBeenCalledWith('v1', undefined)
    registration.dispose()
    expect(disposers.every((dispose) => dispose.mock.calls.length === 1)).toBe(true)
  })

  it('rejects malformed command parameters before invoking the controller', async () => {
    const { commands, controller, univer } = harness()
    registerUniverHistoryCommands(univer, controller)
    const byId = (id: string) => commands.find((entry) => entry.id === id)!
    await expect(byId(INJOFFICE_HISTORY_COMMANDS.preview).handler(undefined, { versionId: '' })).resolves.toBe(false)
    await expect(byId(INJOFFICE_HISTORY_COMMANDS.capture).handler(undefined, { author: { id: 'u1', kind: 'bot' } })).resolves.toBe(false)
    await expect(byId(INJOFFICE_HISTORY_COMMANDS.restore).handler(undefined, { sourceVersionId: 'v1' })).resolves.toBe(false)
    expect(controller.preview).not.toHaveBeenCalled()
    expect(controller.capture).not.toHaveBeenCalled()
    expect(controller.restore).not.toHaveBeenCalled()
  })
})

describe('Univer history UI entry', () => {
  it('auto-opens the host sidebar and registers a Start/History ribbon item', async () => {
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
    const registration = registerUniverHistoryUI(univer, { open }, {
      title: 'Saved versions',
      tooltip: 'Open history',
      order: 7,
      menuHidden$,
    })

    expect(open).toHaveBeenCalledOnce()
    expect(commands.map(({ id }) => id)).toEqual([INJOFFICE_HISTORY_UI_COMMANDS.open])
    await expect(commands[0].handler()).resolves.toBe(true)
    expect(open).toHaveBeenCalledTimes(2)
    const item = menus[0]['ribbon.start']['ribbon.start.history'][INJOFFICE_HISTORY_UI_COMMANDS.open].menuItemFactory()
    expect(item).toMatchObject({ title: 'Saved versions', tooltip: 'Open history' })
    expect(item.hidden$).toBe(menuHidden$)
    registration.dispose()
    expect(dispose).toHaveBeenCalledOnce()
  })

  it('can omit auto-open and contain host failures', async () => {
    const get = vi.fn()
    const univer = { __getInjector: () => ({ get }) } as unknown as Univer
    registerUniverHistoryUI(univer, { open: vi.fn() }, { command: false, autoOpen: false })
    expect(get).not.toHaveBeenCalled()

    const commands: any[] = []
    const failing = { __getInjector: () => ({ get: () => ({ registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } }) }) } as unknown as Univer
    registerUniverHistoryUI(failing, { open: () => { throw new Error('mount failed') } }, { menu: false, autoOpen: false })
    await expect(commands[0].handler()).resolves.toBe(false)
  })
})
