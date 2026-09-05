import { CommandType, ICommandService, IUndoRedoService, type IDisposable, type Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets/lib/facade'
import { IMenuManagerService, MenuItemType, RibbonInsertGroup, RibbonPosition } from '@univerjs/ui'
import type { Observable } from 'rxjs'
import { SparklineCommandController } from './commands'
import { SparklineManager, type CreateSparklineInput } from './manager'
import type { SparklineSnapshotV1, SparklineSpec } from './types'

export const INJOFFICE_SPARKLINE_COMMANDS = {
  restore: 'injoffice.mutation.sparkline.restore',
  create: 'injoffice.command.sparkline.create',
  update: 'injoffice.command.sparkline.update',
  remove: 'injoffice.command.sparkline.remove',
  group: 'injoffice.command.sparkline.group',
  ungroup: 'injoffice.command.sparkline.ungroup',
} as const

export interface UniverSparklineCommandConfig {
  create?: boolean
  update?: boolean
  remove?: boolean
  group?: boolean
  ungroup?: boolean
  menu?: boolean
  /** Runtime menu visibility supplied by a feature composition registry. */
  menuHidden$?: Observable<boolean>
  /** Supplies input when the ribbon button invokes create without params. */
  createInput?: () => CreateSparklineInput | null
}

/**
 * Register lifecycle commands against Univer's public command/undo services.
 * The Insert item is omitted unless a host supplies its selection/dialog input.
 */
export function registerUniverSparklineCommands(
  univer: Univer,
  api: FUniver,
  config: UniverSparklineCommandConfig = {},
  manager = new SparklineManager(),
): { controller: SparklineCommandController; dispose(): void } {
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const controller = new SparklineCommandController(manager, {
    push(record) {
      const unitID = api.getActiveWorkbook()?.getId()
      if (!unitID) return
      undoRedo.pushUndoRedo({
        unitID,
        undoMutations: [{ id: INJOFFICE_SPARKLINE_COMMANDS.restore, params: { snapshot: record.before } }],
        redoMutations: [{ id: INJOFFICE_SPARKLINE_COMMANDS.restore, params: { snapshot: record.after } }],
      })
    },
  })
  const registrations: IDisposable[] = []
  registrations.push(commandService.registerCommand({
    id: INJOFFICE_SPARKLINE_COMMANDS.restore,
    type: CommandType.MUTATION,
    handler: (_accessor, params: { snapshot?: SparklineSnapshotV1 } = {}) =>
      !!params.snapshot && controller.restore(params.snapshot),
  }))
  if (config.create !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SPARKLINE_COMMANDS.create,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { input?: CreateSparklineInput } = {}) => {
      const input = params.input ?? config.createInput?.()
      if (!input) return false
      try { controller.create(input); return true } catch { return false }
    },
  }))
  if (config.update !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SPARKLINE_COMMANDS.update,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; patch?: Partial<Omit<SparklineSpec, 'id' | 'groupId'>> } = {}) => {
      if (typeof params.id !== 'string' || !params.patch) return false
      try { controller.update(params.id, params.patch); return true } catch { return false }
    },
  }))
  if (config.remove !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SPARKLINE_COMMANDS.remove,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string } = {}) => typeof params.id === 'string' && controller.remove(params.id),
  }))
  if (config.group !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SPARKLINE_COMMANDS.group,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { memberIds?: string[]; id?: string } = {}) => {
      if (!Array.isArray(params.memberIds)) return false
      try { controller.group(params.memberIds, params.id); return true } catch { return false }
    },
  }))
  if (config.ungroup !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SPARKLINE_COMMANDS.ungroup,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { memberIds?: string[] } = {}) => {
      if (!Array.isArray(params.memberIds)) return false
      controller.ungroup(params.memberIds)
      return true
    },
  }))

  if (config.menu !== false && config.create !== false && config.createInput) {
    injector.get(IMenuManagerService).mergeMenu({
      [RibbonPosition.INSERT]: {
        [RibbonInsertGroup.MEDIA]: {
          [INJOFFICE_SPARKLINE_COMMANDS.create]: {
            order: 40,
            menuItemFactory: () => ({
              id: INJOFFICE_SPARKLINE_COMMANDS.create,
              title: 'Sparkline',
              tooltip: 'Insert sparkline',
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
