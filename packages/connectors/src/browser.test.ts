import type { ICellData, Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_CONNECTOR_COMMANDS, registerUniverConnectorCommands } from './browser'

function apiHarness(): FUniver {
  const values: ICellData[][] = [[{ v: 'old' }]]
  const sheet = {
    getSheetId: () => 'sheet-1',
    getActiveRange: () => ({ getRange: () => ({ startRow: 0, startColumn: 0 }) }),
    getRange: () => ({
      getValues: () => [[values[0][0].v ?? null]],
      getFormulas: () => [[values[0][0].f ?? null]],
      setValues: (next: typeof values) => { values[0][0] = structuredClone(next[0][0]) },
    }),
  }
  const workbook = { getId: () => 'book-1', getActiveSheet: () => sheet, getActiveRange: () => sheet.getActiveRange(), getSheetBySheetId: () => sheet }
  return { getActiveWorkbook: () => workbook } as unknown as FUniver
}

describe('Univer connector commands', () => {
  it('registers independently configurable commands, Data UI, and atomic snapshot undo', async () => {
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
    let registration!: ReturnType<typeof registerUniverConnectorCommands>
    registration = registerUniverConnectorCommands(univer, apiHarness(), async () => [['new']], {
      add: false,
      refreshMenu: false,
      createInput: () => ({ name: 'Feed', source: { kind: 'http', url: '/gateway/feed', format: 'json' }, refresh: 'manual' }),
      activeConnectorId: () => registration.controller.list()[0]?.id ?? null,
    })
    expect(commands.map(({ id }) => id)).not.toContain(INJOFFICE_CONNECTOR_COMMANDS.add)
    expect(menus).toHaveLength(1)
    expect(JSON.stringify(menus[0])).toContain(INJOFFICE_CONNECTOR_COMMANDS.create)
    expect(JSON.stringify(menus[0])).not.toContain(INJOFFICE_CONNECTOR_COMMANDS.refresh)
    expect(await commands.find(({ id }) => id === INJOFFICE_CONNECTOR_COMMANDS.create)!.handler()).toBe(true)
    expect(registration.controller.list()).toHaveLength(1)
    expect(undos).toMatchObject([{
      unitID: 'book-1',
      undoMutations: [{ id: INJOFFICE_CONNECTOR_COMMANDS.restore, params: { snapshot: { version: 1, manager: { version: 1, connectors: [] } } } }],
    }])
    const restore = commands.find(({ id }) => id === INJOFFICE_CONNECTOR_COMMANDS.restore)!
    expect(restore.handler(undefined, undos[0].undoMutations[0].params)).toBe(true)
    expect(registration.controller.list()).toEqual([])
    registration.dispose()
    expect(disposers.every((dispose) => dispose.mock.calls.length === 1)).toBe(true)
  })

  it('omits UI without host providers and rejects malformed command parameters', async () => {
    const commands: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const services = [
      { registerCommand: (command: any) => { commands.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: vi.fn() },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    expect(() => registerUniverConnectorCommands(univer, apiHarness(), async () => [[1]])).not.toThrow()
    expect(commands.find(({ id }) => id === INJOFFICE_CONNECTOR_COMMANDS.add)?.handler(undefined, {})).toBe(false)
    expect(commands.find(({ id }) => id === INJOFFICE_CONNECTOR_COMMANDS.update)?.handler(undefined, { id: '', patch: {} })).toBe(false)
    expect(await commands.find(({ id }) => id === INJOFFICE_CONNECTOR_COMMANDS.refresh)?.handler(undefined, { id: 'missing' })).toBe(false)
  })
})
