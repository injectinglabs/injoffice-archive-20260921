import type { Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_SHAPE_COMMANDS, registerUniverShapeCommands } from './browser'

function apiHarness(): FUniver {
  const sheet = { getSheetId: () => 'sheet-1', addFloatDomToPosition: (_config: unknown, id: string) => ({ id, dispose: vi.fn() }) }
  return { getActiveWorkbook: () => ({ getId: () => 'book-1', getActiveSheet: () => sheet, getSheetBySheetId: () => sheet }) } as unknown as FUniver
}

describe('Univer shape commands', () => {
  it('registers configurable lifecycle commands, menu, and snapshot undo', () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const undos: any[] = []
    const menus: unknown[] = []
    const disposers: Array<ReturnType<typeof vi.fn>> = []
    const services = [
      { registerCommand: (command: any) => { commands.push(command); const dispose = vi.fn(); disposers.push(dispose); return { dispose } } },
      { pushUndoRedo: (item: any) => undos.push(item) },
      { mergeMenu: (menu: unknown) => menus.push(menu) },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const registration = registerUniverShapeCommands(univer, apiHarness(), { add: false, createKind: () => 'ellipse' })
    expect(commands.map(({ id }) => id)).not.toContain(INJOFFICE_SHAPE_COMMANDS.add)
    expect(menus).toHaveLength(1)
    const create = commands.find(({ id }) => id === INJOFFICE_SHAPE_COMMANDS.create)!
    expect(create.handler()).toBe(true)
    expect(registration.controller.list()[0].value?.kind).toBe('ellipse')
    expect(undos).toMatchObject([{
      unitID: 'book-1',
      undoMutations: [{ id: INJOFFICE_SHAPE_COMMANDS.restore, params: { snapshot: { version: 1, shapes: [] } } }],
    }])
    const restore = commands.find(({ id }) => id === INJOFFICE_SHAPE_COMMANDS.restore)!
    expect(restore.handler(undefined, undos[0].undoMutations[0].params)).toBe(true)
    expect(registration.controller.list()).toEqual([])
    registration.dispose()
    expect(disposers.every((dispose) => dispose.mock.calls.length === 1)).toBe(true)
  })

  it('omits Insert UI without a provider and rejects malformed params', () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const services = [
      { registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: vi.fn() },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    expect(() => registerUniverShapeCommands(univer, apiHarness())).not.toThrow()
    expect(commands.find(({ id }) => id === INJOFFICE_SHAPE_COMMANDS.create)?.handler(undefined, { kind: 'unknown' })).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_SHAPE_COMMANDS.update)?.handler(undefined, {})).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_SHAPE_COMMANDS.remove)?.handler(undefined, {})).toBe(false)
  })
})
