import type { Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_PIVOT_COMMANDS, registerUniverPivotCommands } from './browser'

const GRID = [
  ['Region', 'Quarter', 'Sales'],
  ['West', 'Q1', 20],
  ['East', 'Q2', 35],
]

function apiHarness() {
  const sheet = {
    getSheetId: () => 'sheet-1',
    getActiveRange: () => ({ getRange: () => ({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 }) }),
    getRange: () => ({ getValues: () => GRID, setValues: vi.fn() }),
  }
  return { getActiveWorkbook: () => ({
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getActiveRange: () => null,
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }) } as unknown as FUniver
}

describe('Univer pivot commands', () => {
  it('registers independently configurable commands, Insert UI, and snapshot undo', () => {
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
    const registration = registerUniverPivotCommands(univer, apiHarness(), {
      add: false, update: false, remove: false, setColumns: false, setValues: false,
      setFilters: false, setMemberFilters: false, setPageFields: false, setSorts: false,
    })
    expect(commands.map(({ id }) => id).sort()).toEqual([
      INJOFFICE_PIVOT_COMMANDS.restore,
      INJOFFICE_PIVOT_COMMANDS.create,
      INJOFFICE_PIVOT_COMMANDS.setRows,
    ].sort())
    expect(menus).toHaveLength(1)
    expect(commands.find(({ id }) => id === INJOFFICE_PIVOT_COMMANDS.create)?.handler()).toBe(true)
    const id = registration.controller.list()[0].id
    expect(commands.find(({ id: commandId }) => commandId === INJOFFICE_PIVOT_COMMANDS.setRows)?.handler(undefined, { id, rows: ['Quarter'] })).toBe(true)
    expect(undos).toHaveLength(2)
    expect(undos[1]).toMatchObject({
      unitID: 'book-1',
      undoMutations: [{ id: INJOFFICE_PIVOT_COMMANDS.restore, params: { snapshot: { version: 1 } } }],
      redoMutations: [{ id: INJOFFICE_PIVOT_COMMANDS.restore, params: { snapshot: { version: 1 } } }],
    })
    const restore = commands.find(({ id: commandId }) => commandId === INJOFFICE_PIVOT_COMMANDS.restore)!
    expect(restore.handler(undefined, undos[1].undoMutations[0].params)).toBe(true)
    expect(registration.controller.getById(id)?.value?.rows).toEqual(['Region'])
    registration.dispose()
    expect(disposers.every((dispose) => dispose.mock.calls.length === 1)).toBe(true)
  })

  it('omits the menu and rejects malformed command parameters', () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const services = [
      { registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: vi.fn() },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const registration = registerUniverPivotCommands(univer, apiHarness(), { menu: false })
    expect(commands.find(({ id }) => id === INJOFFICE_PIVOT_COMMANDS.setValues)?.handler(undefined, { id: 'missing', values: 'bad' })).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_PIVOT_COMMANDS.setFilters)?.handler(undefined, { id: 'missing' })).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_PIVOT_COMMANDS.update)?.handler(undefined, {})).toBe(false)
    expect(registration.controller.list()).toEqual([])
  })
})
