import { CommandType, ICommandService, IUndoRedoService, type IDisposable, type Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets/lib/facade'
import { IMenuManagerService, MenuItemType, RibbonInsertGroup, RibbonPosition } from '@univerjs/ui'
import type { Observable } from 'rxjs'
import { ShapeCommandController, type ShapeSnapshotV1 } from './commands'
import type { FileShapeAnchor } from './fromFile'
import { ShapeManager } from './manager'
import { SHAPE_KINDS, type ShapeKind, type ShapeSpec } from './types'

export * from './index'

export const INJOFFICE_SHAPE_COMMANDS = {
  restore: 'injoffice.mutation.shape.restore',
  create: 'injoffice.command.shape.create',
  add: 'injoffice.command.shape.add',
  update: 'injoffice.command.shape.update',
  remove: 'injoffice.command.shape.remove',
} as const

export interface UniverShapeCommandConfig {
  create?: boolean
  add?: boolean
  update?: boolean
  remove?: boolean
  menu?: boolean
  /** Runtime menu visibility supplied by a feature composition registry. */
  menuHidden$?: Observable<boolean>
  /** Supplies a shape kind when the Insert ribbon invokes create. */
  createKind?: () => ShapeKind | null
}

const KINDS: ReadonlySet<string> = new Set(SHAPE_KINDS)

/** Register shape lifecycle commands against Univer's public command and undo
 * services. The Insert item is omitted unless the host supplies picker input. */
export function registerUniverShapeCommands(
  univer: Univer,
  api: FUniver,
  config: UniverShapeCommandConfig = {},
  manager = new ShapeManager(api),
): { controller: ShapeCommandController; dispose(): void } {
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const controller = new ShapeCommandController(manager, {
    push(record) {
      const unitID = api.getActiveWorkbook()?.getId()
      if (!unitID) return
      undoRedo.pushUndoRedo({
        unitID,
        undoMutations: [{ id: INJOFFICE_SHAPE_COMMANDS.restore, params: { snapshot: record.before } }],
        redoMutations: [{ id: INJOFFICE_SHAPE_COMMANDS.restore, params: { snapshot: record.after } }],
      })
    },
  })
  const registrations: IDisposable[] = []
  registrations.push(commandService.registerCommand({
    id: INJOFFICE_SHAPE_COMMANDS.restore,
    type: CommandType.MUTATION,
    handler: (_accessor, params: { snapshot?: ShapeSnapshotV1 } = {}) => !!params.snapshot && controller.restore(params.snapshot),
  }))
  if (config.create !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SHAPE_COMMANDS.create,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { kind?: unknown } = {}) => {
      const kind = params.kind ?? config.createKind?.()
      return typeof kind === 'string' && KINDS.has(kind) && controller.create(kind as ShapeKind) !== null
    },
  }))
  if (config.add !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SHAPE_COMMANDS.add,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { spec?: ShapeSpec; cellAnchor?: FileShapeAnchor; sheetId?: string } = {}) =>
      !!params.spec && controller.add(params.spec, params.cellAnchor, params.sheetId) !== null,
  }))
  if (config.update !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SHAPE_COMMANDS.update,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; patch?: Partial<Omit<ShapeSpec, 'id' | 'kind' | 'nativeIdentity'>> } = {}) =>
      typeof params.id === 'string' && !!params.patch && controller.update(params.id, params.patch),
  }))
  if (config.remove !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_SHAPE_COMMANDS.remove,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string } = {}) => typeof params.id === 'string' && controller.remove(params.id),
  }))

  if (config.menu !== false && config.create !== false && config.createKind) {
    injector.get(IMenuManagerService).mergeMenu({
      [RibbonPosition.INSERT]: {
        [RibbonInsertGroup.MEDIA]: {
          [INJOFFICE_SHAPE_COMMANDS.create]: {
            order: 50,
            menuItemFactory: () => ({ id: INJOFFICE_SHAPE_COMMANDS.create, title: 'Shape', tooltip: 'Insert shape', type: MenuItemType.BUTTON, hidden$: config.menuHidden$ }),
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
