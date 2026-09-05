import type { Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_SPARKLINE_COMMANDS, registerUniverSparklineCommands } from './browser'

const input = {
  id: 'spark',
  type: 'line' as const,
  source: { sheetId: 's', startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
  target: { sheetId: 's', row: 0, column: 3 },
}

describe('Univer sparkline commands', () => {
  it('registers configurable lifecycle commands and snapshot undo', () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const undos: any[] = []
    const menus: unknown[] = []
    const services = [
      { registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: (item: any) => undos.push(item) },
      { mergeMenu: (menu: unknown) => menus.push(menu) },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const api = { getActiveWorkbook: () => ({ getId: () => 'book' }) } as unknown as FUniver
    const registration = registerUniverSparklineCommands(univer, api, { update: false, createInput: () => input })

    expect(commands.map(({ id }) => id)).not.toContain(INJOFFICE_SPARKLINE_COMMANDS.update)
    expect(menus).toHaveLength(1)
    const create = commands.find(({ id }) => id === INJOFFICE_SPARKLINE_COMMANDS.create)!
    expect(create.handler()).toBe(true)
    expect(registration.controller.getById('spark')?.value).toMatchObject({ id: 'spark' })
    expect(undos).toMatchObject([{
      unitID: 'book',
      undoMutations: [{ id: INJOFFICE_SPARKLINE_COMMANDS.restore, params: { snapshot: { sparklines: [] } } }],
    }])
    const restore = commands.find(({ id }) => id === INJOFFICE_SPARKLINE_COMMANDS.restore)!
    expect(restore.handler(undefined, undos[0].undoMutations[0].params)).toBe(true)
    expect(registration.controller.list()).toEqual([])
  })

  it('omits Insert UI without a host input provider', () => {
    const services = [
      { registerCommand: () => ({ dispose: vi.fn() }) },
      { pushUndoRedo: vi.fn() },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const api = { getActiveWorkbook: () => ({ getId: () => 'book' }) } as unknown as FUniver
    expect(() => registerUniverSparklineCommands(univer, api)).not.toThrow()
  })
})
