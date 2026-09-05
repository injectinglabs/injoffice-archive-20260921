import { PrintManager } from './manager'
import type { PrintConfigurationSnapshotV1, PrintLayoutConfig, PrintRenderConfig } from './types'
import { validatePrintConfigurationSnapshot } from './validation'

export interface PrintConfigurationUndoRecord {
  label: string
  before: PrintConfigurationSnapshotV1
  after: PrintConfigurationSnapshotV1
}

export interface PrintConfigurationUndoSink {
  push(record: PrintConfigurationUndoRecord): void
}

export interface PrintConfigurationPersistence {
  /** Save one complete validated replacement. Implementations must reject the
   * promise when the durable write was not committed. */
  save(snapshot: Readonly<PrintConfigurationSnapshotV1>): Promise<void> | void
}

export interface PrintConfigurationCommandTarget {
  updateLayout(patch: Partial<PrintLayoutConfig>): Promise<boolean> | boolean
  updateRender(patch: Partial<PrintRenderConfig>): Promise<boolean> | boolean
  replace(snapshot: unknown): Promise<boolean> | boolean
  restore(snapshot: unknown): Promise<boolean> | boolean
  dispose?(): void
}

export interface PrintAuthoritativeApplyOptions {
  label?: string
  recordUndo?: boolean
}

export class PrintConfigurationPersistenceError extends Error {
  constructor(cause: unknown) {
    super('print configuration persistence failed', { cause })
    this.name = 'PrintConfigurationPersistenceError'
  }
}

function immutableCopy(snapshot: PrintConfigurationSnapshotV1): Readonly<PrintConfigurationSnapshotV1> {
  const copy = structuredClone(snapshot)
  const seen = new WeakSet<object>()
  const freeze = (value: unknown): void => {
    if (!value || typeof value !== 'object' || seen.has(value)) return
    seen.add(value)
    for (const child of Object.values(value)) freeze(child)
    Object.freeze(value)
  }
  freeze(copy)
  return copy
}

/** Serialized, persistence-first command lifecycle. Durable-save rejection
 * leaves local state and undo history untouched. */
export class PrintConfigurationCommandController {
  private transition: Promise<void> = Promise.resolve()

  constructor(
    readonly manager: PrintManager,
    private readonly persistence: PrintConfigurationPersistence,
    private readonly undo?: PrintConfigurationUndoSink,
  ) {
    if (!persistence || typeof persistence.save !== 'function') throw new TypeError('persistence.save must be a function')
  }

  snapshot(): PrintConfigurationSnapshotV1 {
    return this.manager.configurationSnapshot()
  }

  updateLayout(patch: Partial<PrintLayoutConfig>): Promise<boolean> {
    return this.enqueue(async () => {
      if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return false
      let candidate: PrintConfigurationSnapshotV1
      try {
        const before = this.snapshot()
        candidate = { ...before, layout: { ...before.layout, ...structuredClone(patch) } }
      } catch {
        return false
      }
      return this.commit('Update print layout', candidate, true)
    })
  }

  updateRender(patch: Partial<PrintRenderConfig>): Promise<boolean> {
    return this.enqueue(async () => {
      if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return false
      let candidate: PrintConfigurationSnapshotV1
      try {
        const before = this.snapshot()
        candidate = {
          ...before,
          render: {
            ...before.render,
            ...structuredClone(patch),
            headerFooterSetting: { ...before.render.headerFooterSetting, ...structuredClone(patch?.headerFooterSetting ?? {}) },
          },
        }
      } catch {
        return false
      }
      return this.commit('Update print rendering', candidate, true)
    })
  }

  replace(snapshot: unknown): Promise<boolean> {
    return this.enqueue(() => this.commit('Replace print configuration', snapshot, true))
  }

  /** Undo/redo restore persists but never creates another undo record. */
  restore(snapshot: unknown): Promise<boolean> {
    return this.enqueue(() => this.commit('Restore print configuration', snapshot, false))
  }

  /** Apply one already-authorized complete server state through the same
   * persistence-first critical section as local commands. Collaboration can
   * record an acknowledged local command or omit remote/resync history. */
  applyAuthoritative(snapshot: unknown, options: PrintAuthoritativeApplyOptions = {}): Promise<boolean> {
    const label = typeof options.label === 'string' && options.label.trim() && options.label.length <= 256
      ? options.label
      : 'Apply print configuration'
    return this.enqueue(() => this.commit(label, snapshot, options.recordUndo === true))
  }

  private enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const next = this.transition.then(operation, operation)
    this.transition = next.then(() => undefined, () => undefined)
    return next
  }

  private async commit(label: string, input: unknown, recordUndo: boolean): Promise<boolean> {
    const checked = validatePrintConfigurationSnapshot(input)
    if (!checked.ok) return false
    const before = this.snapshot()
    const after = checked.value
    if (JSON.stringify(before) === JSON.stringify(after)) return true
    try {
      await this.persistence.save(immutableCopy(after))
    } catch (cause) {
      throw new PrintConfigurationPersistenceError(cause)
    }
    if (!this.manager.restoreConfiguration(after)) throw new Error('validated print configuration could not be applied')
    if (recordUndo) this.undo?.push({ label, before, after: this.snapshot() })
    return true
  }
}
