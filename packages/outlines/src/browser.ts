import {
  CommandType,
  ICommandService,
  IUndoRedoService,
  type IDisposable,
  type Univer,
} from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets/lib/facade'
import { IMenuManagerService, MenuItemType, RibbonDataGroup, RibbonPosition } from '@univerjs/ui'
import type { OutlineCollaborationCommandTarget } from './collaboration'
import type { Observable } from 'rxjs'
import { OutlineCommandController, type OutlineUndoSink } from './commands'
import { OutlineManager } from './manager'
import type { OutlineAxis, OutlineGroup, OutlineResult, OutlineVisibilityAdapter } from './types'

export * from './index'

export function createUniverOutlineVisibilityAdapter(api: FUniver): OutlineVisibilityAdapter {
  return {
    hide(sheetId, axis, start, count) {
      const sheet = api.getActiveWorkbook()?.getSheetBySheetId(sheetId)
      if (!sheet) return
      if (axis === 'row') sheet.hideRows(start, count)
      else sheet.hideColumns(start, count)
    },
    show(sheetId, axis, start, count) {
      const sheet = api.getActiveWorkbook()?.getSheetBySheetId(sheetId)
      if (!sheet) return
      if (axis === 'row') sheet.showRows(start, count)
      else sheet.showColumns(start, count)
    },
  }
}

let sequence = 0

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

/** Selection-aware facade for wiring group/ungroup commands into a Univer host. */
export class UniverOutlineController {
  readonly manager: OutlineManager
  readonly commands: OutlineCommandController
  private readonly api: FUniver

  constructor(
    api: FUniver,
    manager = new OutlineManager(createUniverOutlineVisibilityAdapter(api)),
    undo?: OutlineUndoSink,
  ) {
    this.api = api
    this.manager = manager
    this.commands = new OutlineCommandController(manager, undo)
  }

  addFromSelection(axis: OutlineAxis, collapsed = false, id?: string): OutlineResult<OutlineGroup> | null {
    const workbook = this.api.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    const range = sheet?.getActiveRange() ?? workbook?.getActiveRange()
    if (!sheet || !range) return null
    const selected = range.getRange()
    return this.commands.add({
      id: id ?? `outline-${Date.now().toString(36)}-${(++sequence).toString(36)}`,
      sheetId: sheet.getSheetId(),
      axis,
      start: axis === 'row' ? selected.startRow : selected.startColumn,
      end: axis === 'row' ? selected.endRow : selected.endColumn,
      collapsed,
    })
  }

  clearSelection(axis: OutlineAxis): OutlineGroup[] {
    const workbook = this.api.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    const range = sheet?.getActiveRange() ?? workbook?.getActiveRange()
    if (!sheet || !range) return []
    const selected = range.getRange()
    return this.commands.clear(
      sheet.getSheetId(),
      axis,
      axis === 'row' ? selected.startRow : selected.startColumn,
      axis === 'row' ? selected.endRow : selected.endColumn,
    )
  }
}

export const INJOFFICE_OUTLINE_COMMANDS = {
  restore: 'injoffice.mutation.outline.restore',
  groupRows: 'injoffice.command.outline.group-rows',
  groupColumns: 'injoffice.command.outline.group-columns',
  ungroupRows: 'injoffice.command.outline.ungroup-rows',
  ungroupColumns: 'injoffice.command.outline.ungroup-columns',
  setCollapsed: 'injoffice.command.outline.set-collapsed',
  update: 'injoffice.command.outline.update',
  remove: 'injoffice.command.outline.remove',
} as const

export interface UniverOutlineFeatureConfig {
  groupRows?: boolean
  groupColumns?: boolean
  ungroupRows?: boolean
  ungroupColumns?: boolean
  setCollapsed?: boolean
  update?: boolean
  remove?: boolean
  menu?: boolean
  /** Route registered commands and undo/redo restore through a server-first
   * collaboration target. The registration owns and disposes that target. */
  commandTargetFactory?: (controller: OutlineCommandController) => OutlineCollaborationCommandTarget
  /** Runtime menu visibility supplied by a feature composition registry. */
  menuHidden$?: Observable<boolean>
}

/** Register public Univer commands and optional Data-ribbon items. */
export function registerUniverOutlineCommands(
  univer: Univer,
  api: FUniver,
  config: UniverOutlineFeatureConfig = {},
  manager = new OutlineManager(createUniverOutlineVisibilityAdapter(api)),
): { controller: UniverOutlineController; dispose(): void } {
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const unitId = () => api.getActiveWorkbook()?.getId()
  const controller = new UniverOutlineController(api, manager, {
    push(record) {
      const currentUnitId = unitId()
      if (!currentUnitId) return
      undoRedo.pushUndoRedo({
        unitID: currentUnitId,
        undoMutations: [{ id: INJOFFICE_OUTLINE_COMMANDS.restore, params: { snapshot: record.before } }],
        redoMutations: [{ id: INJOFFICE_OUTLINE_COMMANDS.restore, params: { snapshot: record.after } }],
      })
    },
  })
  const localTarget: OutlineCollaborationCommandTarget = {
    add: (group) => controller.commands.add(group).ok,
    update: (id, patch) => controller.commands.update(id, patch).ok,
    setCollapsed: (id, collapsed) => controller.commands.setCollapsed(id, collapsed).ok,
    remove: (id) => controller.commands.remove(id),
    clear: (sheetId, axis, start, end) => controller.commands.clear(sheetId, axis, start, end).length > 0,
    restore: (snapshot) => controller.commands.restore(snapshot),
  }
  const target = config.commandTargetFactory?.(controller.commands) ?? localTarget
  if (!target || ['add', 'update', 'setCollapsed', 'remove', 'clear', 'restore'].some((name) => typeof target[name as keyof OutlineCollaborationCommandTarget] !== 'function')) {
    throw new TypeError('commandTargetFactory must return an outline collaboration command target')
  }
  const registrations: IDisposable[] = []
  const safe = async (operation: () => boolean | Promise<boolean>): Promise<boolean> => {
    try { return await operation() === true } catch { return false }
  }
  const selectionGroup = (axis: OutlineAxis): OutlineGroup | null => {
    const workbook = api.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    const range = sheet?.getActiveRange() ?? workbook?.getActiveRange()
    if (!sheet || !range) return null
    const selected = range.getRange()
    return {
      id: `outline-${Date.now().toString(36)}-${(++sequence).toString(36)}`,
      sheetId: sheet.getSheetId(),
      axis,
      start: axis === 'row' ? selected.startRow : selected.startColumn,
      end: axis === 'row' ? selected.endRow : selected.endColumn,
      collapsed: false,
    }
  }
  const clearSelection = (axis: OutlineAxis): Promise<boolean> => {
    const workbook = api.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    const range = sheet?.getActiveRange() ?? workbook?.getActiveRange()
    if (!sheet || !range) return Promise.resolve(false)
    const selected = range.getRange()
    return safe(() => target.clear(
      sheet.getSheetId(),
      axis,
      axis === 'row' ? selected.startRow : selected.startColumn,
      axis === 'row' ? selected.endRow : selected.endColumn,
    ))
  }
  registrations.push(commandService.registerCommand({
    id: INJOFFICE_OUTLINE_COMMANDS.restore,
    type: CommandType.MUTATION,
    handler: (_accessor, params: unknown = {}) => {
      if (!isRecord(params) || !Array.isArray(params.snapshot)) return false
      return safe(() => target.restore(params.snapshot as OutlineGroup[]))
    },
  }))

  const actions = [
    ['groupRows', INJOFFICE_OUTLINE_COMMANDS.groupRows, () => { const group = selectionGroup('row'); return group ? safe(() => target.add(group)) : false }],
    ['groupColumns', INJOFFICE_OUTLINE_COMMANDS.groupColumns, () => { const group = selectionGroup('column'); return group ? safe(() => target.add(group)) : false }],
    ['ungroupRows', INJOFFICE_OUTLINE_COMMANDS.ungroupRows, () => clearSelection('row')],
    ['ungroupColumns', INJOFFICE_OUTLINE_COMMANDS.ungroupColumns, () => clearSelection('column')],
  ] as const
  for (const [key, id, handler] of actions) {
    if (config[key] === false) continue
    registrations.push(commandService.registerCommand({ id, type: CommandType.COMMAND, handler }))
  }
  if (config.setCollapsed !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_OUTLINE_COMMANDS.setCollapsed,
    type: CommandType.COMMAND,
    handler: (_accessor, params: unknown = {}) => {
      if (!isRecord(params) || typeof params.id !== 'string' || typeof params.collapsed !== 'boolean') return false
      return safe(() => target.setCollapsed(params.id as string, params.collapsed as boolean))
    },
  }))
  if (config.update !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_OUTLINE_COMMANDS.update,
    type: CommandType.COMMAND,
    handler: (_accessor, params: unknown = {}) => {
      if (!isRecord(params) || typeof params.id !== 'string' || !isRecord(params.patch)) return false
      return safe(() => target.update(params.id as string, params.patch as Partial<Omit<OutlineGroup, 'id'>>))
    },
  }))
  if (config.remove !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_OUTLINE_COMMANDS.remove,
    type: CommandType.COMMAND,
    handler: (_accessor, params: unknown = {}) => {
      if (!isRecord(params) || typeof params.id !== 'string') return false
      return safe(() => target.remove(params.id as string))
    },
  }))

  if (config.menu !== false) {
    const menu = injector.get(IMenuManagerService)
    const entries = [
      ['groupRows', INJOFFICE_OUTLINE_COMMANDS.groupRows, 'Group rows', 10],
      ['groupColumns', INJOFFICE_OUTLINE_COMMANDS.groupColumns, 'Group columns', 20],
      ['ungroupRows', INJOFFICE_OUTLINE_COMMANDS.ungroupRows, 'Ungroup rows', 30],
      ['ungroupColumns', INJOFFICE_OUTLINE_COMMANDS.ungroupColumns, 'Ungroup columns', 40],
    ] as const
    for (const [key, id, title, order] of entries) {
      if (config[key] === false) continue
      menu.mergeMenu({
        [RibbonPosition.DATA]: {
          [RibbonDataGroup.ORGANIZATION]: {
            [id]: { order, menuItemFactory: () => ({ id, title, tooltip: title, type: MenuItemType.BUTTON, hidden$: config.menuHidden$ }) },
          },
        },
      })
    }
  }
  return {
    controller,
    dispose() {
      for (const registration of registrations.reverse()) registration.dispose()
      if (target !== localTarget) target.dispose?.()
    },
  }
}
