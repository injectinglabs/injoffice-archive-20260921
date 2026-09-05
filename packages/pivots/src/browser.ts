import { CommandType, ICommandService, IUndoRedoService, type IDisposable, type Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets/lib/facade'
import { IMenuManagerService, MenuItemType, RibbonInsertGroup, RibbonPosition } from '@univerjs/ui'
import type { Observable } from 'rxjs'
import { PivotCommandController, type PivotSnapshotV1 } from './commands'
import { PivotManager } from './manager'
import type { PivotFieldSort, PivotMemberFilter, PivotPageField, PivotSpec, PivotValueField } from './types'

export * from './index'

export const INJOFFICE_PIVOT_COMMANDS = {
  restore: 'injoffice.mutation.pivot.restore',
  create: 'injoffice.command.pivot.create',
  add: 'injoffice.command.pivot.add',
  update: 'injoffice.command.pivot.update',
  remove: 'injoffice.command.pivot.remove',
  setRows: 'injoffice.command.pivot.set-rows',
  setColumns: 'injoffice.command.pivot.set-columns',
  setValues: 'injoffice.command.pivot.set-values',
  setFilters: 'injoffice.command.pivot.set-filters',
  setMemberFilters: 'injoffice.command.pivot.set-member-filters',
  setPageFields: 'injoffice.command.pivot.set-page-fields',
  setSorts: 'injoffice.command.pivot.set-sorts',
} as const

export interface UniverPivotCommandConfig {
  create?: boolean
  add?: boolean
  update?: boolean
  remove?: boolean
  setRows?: boolean
  setColumns?: boolean
  setValues?: boolean
  setFilters?: boolean
  setMemberFilters?: boolean
  setPageFields?: boolean
  setSorts?: boolean
  menu?: boolean
  /** Runtime menu visibility supplied by a feature composition registry. */
  menuHidden$?: Observable<boolean>
}

/** Register configurable pivot commands and complete-snapshot undo/redo. */
export function registerUniverPivotCommands(
  univer: Univer,
  api: FUniver,
  config: UniverPivotCommandConfig = {},
  manager = new PivotManager(api),
): { controller: PivotCommandController; dispose(): void } {
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const controller = new PivotCommandController(manager, {
    push(record) {
      const unitID = api.getActiveWorkbook()?.getId()
      if (!unitID) return
      undoRedo.pushUndoRedo({
        unitID,
        undoMutations: [{ id: INJOFFICE_PIVOT_COMMANDS.restore, params: { snapshot: record.before } }],
        redoMutations: [{ id: INJOFFICE_PIVOT_COMMANDS.restore, params: { snapshot: record.after } }],
      })
    },
  })
  const registrations: IDisposable[] = []
  registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.restore,
    type: CommandType.MUTATION,
    handler: (_accessor, params: { snapshot?: PivotSnapshotV1 } = {}) => !!params.snapshot && controller.restore(params.snapshot),
  }))
  if (config.create !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.create,
    type: CommandType.COMMAND,
    handler: () => controller.createFromSelection() !== null,
  }))
  if (config.add !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.add,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { spec?: PivotSpec } = {}) => !!params.spec && controller.add(params.spec) !== null,
  }))
  if (config.update !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.update,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; patch?: Partial<Omit<PivotSpec, 'id' | 'nativeIdentity'>> } = {}) =>
      typeof params.id === 'string' && !!params.patch && controller.update(params.id, params.patch),
  }))
  if (config.remove !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.remove,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string } = {}) => typeof params.id === 'string' && controller.remove(params.id),
  }))
  if (config.setRows !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.setRows,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; rows?: string[] } = {}) => typeof params.id === 'string' && Array.isArray(params.rows) && controller.setRowFields(params.id, params.rows),
  }))
  if (config.setColumns !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.setColumns,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; columns?: string[] } = {}) => typeof params.id === 'string' && Array.isArray(params.columns) && controller.setColumnFields(params.id, params.columns),
  }))
  if (config.setValues !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.setValues,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; values?: PivotValueField[] } = {}) => typeof params.id === 'string' && Array.isArray(params.values) && controller.setValueFields(params.id, params.values),
  }))
  if (config.setFilters !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.setFilters,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; filters?: Record<string, string[]> | null } = {}) =>
      typeof params.id === 'string' && params.filters !== undefined && controller.setFilters(params.id, params.filters ?? undefined),
  }))
  if (config.setMemberFilters !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.setMemberFilters,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; memberFilters?: PivotMemberFilter[] | null } = {}) =>
      typeof params.id === 'string' && params.memberFilters !== undefined && controller.setMemberFilters(params.id, params.memberFilters ?? undefined),
  }))
  if (config.setPageFields !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.setPageFields,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; pageFields?: PivotPageField[] | null } = {}) =>
      typeof params.id === 'string' && params.pageFields !== undefined && controller.setPageFields(params.id, params.pageFields ?? undefined),
  }))
  if (config.setSorts !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_PIVOT_COMMANDS.setSorts,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; sorts?: PivotFieldSort[] | null } = {}) =>
      typeof params.id === 'string' && params.sorts !== undefined && controller.setSorts(params.id, params.sorts ?? undefined),
  }))

  if (config.menu !== false && config.create !== false) {
    injector.get(IMenuManagerService).mergeMenu({
      [RibbonPosition.INSERT]: {
        [RibbonInsertGroup.MEDIA]: {
          [INJOFFICE_PIVOT_COMMANDS.create]: {
            order: 20,
            menuItemFactory: () => ({
              id: INJOFFICE_PIVOT_COMMANDS.create,
              title: 'Pivot table',
              tooltip: 'Insert pivot table from selection',
              type: MenuItemType.BUTTON,
              hidden$: config.menuHidden$,
            }),
          },
        },
      },
    })
  }
  return {
    controller,
    dispose() { for (const registration of registrations.reverse()) registration.dispose() },
  }
}
