import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it, vi } from 'vitest'
import type { Univer } from '@univerjs/core'
import { INJOFFICE_OUTLINE_COMMANDS, UniverOutlineController, registerUniverOutlineCommands } from './browser'

function harness() {
  const hideRows = vi.fn()
  const showRows = vi.fn()
  const hideColumns = vi.fn()
  const showColumns = vi.fn()
  const selected = { startRow: 2, endRow: 5, startColumn: 3, endColumn: 7 }
  const sheet = {
    getSheetId: () => 'sheet-1',
    getActiveRange: () => ({ getRange: () => selected }),
    hideRows,
    showRows,
    hideColumns,
    showColumns,
  }
  const workbook = {
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getActiveRange: () => null,
    getSheetBySheetId: () => sheet,
  }
  const api = { getActiveWorkbook: () => workbook } as unknown as FUniver
  return { api, hideRows, hideColumns }
}

describe('Univer outline controller', () => {
  it('creates collapsed groups from row and column selections', () => {
    const subject = harness()
    const controller = new UniverOutlineController(subject.api)
    expect(controller.addFromSelection('row', true, 'rows')).toMatchObject({
      ok: true,
      value: { id: 'rows', axis: 'row', start: 2, end: 5 },
    })
    expect(controller.addFromSelection('column', true, 'columns')).toMatchObject({
      ok: true,
      value: { id: 'columns', axis: 'column', start: 3, end: 7 },
    })
    expect(subject.hideRows).toHaveBeenCalledWith(2, 4)
    expect(subject.hideColumns).toHaveBeenCalledWith(3, 5)
  })

  it('registers configurable public commands and Univer undo mutations', async () => {
    const subject = harness()
    const registered: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const undoItems: any[] = []
    const menus: unknown[] = []
    const services = [
      { registerCommand: (command: any) => { registered.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: (item: any) => undoItems.push(item) },
      { mergeMenu: (menu: unknown) => menus.push(menu) },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const registration = registerUniverOutlineCommands(univer, subject.api, { groupColumns: false, ungroupColumns: false })

    expect(registered.map(({ id }) => id)).not.toContain(INJOFFICE_OUTLINE_COMMANDS.groupColumns)
    expect(menus).toHaveLength(2)
    const group = registered.find(({ id }) => id === INJOFFICE_OUTLINE_COMMANDS.groupRows)!
    await expect(group.handler()).resolves.toBe(true)
    expect(undoItems).toHaveLength(1)
    expect(undoItems[0]).toMatchObject({
      unitID: 'book-1',
      undoMutations: [{ id: INJOFFICE_OUTLINE_COMMANDS.restore, params: { snapshot: [] } }],
    })
    const restore = registered.find(({ id }) => id === INJOFFICE_OUTLINE_COMMANDS.restore)!
    await expect(restore.handler(undefined, undoItems[0].undoMutations[0].params)).resolves.toBe(true)
    expect(registration.controller.manager.list()).toEqual([])
  })

  it('routes every registered mutation through a disposable collaboration target', async () => {
    const subject = harness()
    const registered: Array<{ id: string; handler: (accessor?: unknown, params?: any) => unknown }> = []
    const services = [
      { registerCommand: (command: any) => { registered.push(command); return { dispose: vi.fn() } } },
      { pushUndoRedo: vi.fn() },
    ]
    let service = 0
    const univer = { __getInjector: () => ({ get: () => services[service++] }) } as unknown as Univer
    const target = {
      add: vi.fn(async () => true),
      update: vi.fn(async () => true),
      setCollapsed: vi.fn(async () => true),
      remove: vi.fn(async () => true),
      clear: vi.fn(async () => true),
      restore: vi.fn(async () => true),
      dispose: vi.fn(),
    }
    const registration = registerUniverOutlineCommands(univer, subject.api, { menu: false, commandTargetFactory: () => target })
    const invoke = (id: string, params?: unknown) => registered.find((command) => command.id === id)!.handler(undefined, params)

    await expect(invoke(INJOFFICE_OUTLINE_COMMANDS.groupRows)).resolves.toBe(true)
    await expect(invoke(INJOFFICE_OUTLINE_COMMANDS.ungroupRows)).resolves.toBe(true)
    await expect(invoke(INJOFFICE_OUTLINE_COMMANDS.setCollapsed, { id: 'rows', collapsed: true })).resolves.toBe(true)
    await expect(invoke(INJOFFICE_OUTLINE_COMMANDS.update, { id: 'rows', patch: { start: 1 } })).resolves.toBe(true)
    await expect(invoke(INJOFFICE_OUTLINE_COMMANDS.remove, { id: 'rows' })).resolves.toBe(true)
    await expect(invoke(INJOFFICE_OUTLINE_COMMANDS.restore, { snapshot: [] })).resolves.toBe(true)
    expect(target.add).toHaveBeenCalledOnce()
    expect(target.clear).toHaveBeenCalledOnce()
    expect(target.setCollapsed).toHaveBeenCalledOnce()
    expect(target.update).toHaveBeenCalledOnce()
    expect(target.remove).toHaveBeenCalledOnce()
    expect(target.restore).toHaveBeenCalledOnce()
    registration.dispose()
    expect(target.dispose).toHaveBeenCalledOnce()
  })
})
