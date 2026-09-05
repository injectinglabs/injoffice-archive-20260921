import { HistoryManager } from './manager'
import type {
  HistoryAuthor,
  HistoryCaptureOptions,
  HistoryListOptions,
  HistoryListPage,
  HistoryPreview,
  HistoryRestoreOptions,
  HistoryVersionInfo,
} from './types'

export type HistoryCommandOperation = 'list' | 'preview' | 'capture' | 'restore'
export type HistoryCommandStatus = 'idle' | 'running' | 'succeeded' | 'failed' | 'cancelled'

export interface HistoryCommandState {
  jobId: string | null
  operation: HistoryCommandOperation | null
  status: HistoryCommandStatus
  error?: string
}

export interface HistoryWorkspace<TSnapshot> {
  captureSnapshot(context: { signal: AbortSignal }): Promise<{ snapshot: TSnapshot; contentType: string }>
  showIsolatedPreview(preview: HistoryPreview<TSnapshot>, context: { signal: AbortSignal }): Promise<void> | void
  /** Optional pre-commit gate used to quiesce collaboration and refuse a
   * restore while local/shared edits are not durably saved. */
  prepareRestore?(context: {
    sourceVersionId: string
    expectedHeadVersionId?: string
    signal: AbortSignal
  }): Promise<HistoryRestorePreparation>
  /** Reload/rejoin the newly created durable version. Called without the
   * cancelled request signal after creation, so the live editor cannot remain
   * on a stale head merely because cancellation raced the durable commit. */
  activateRestoredVersion(version: HistoryVersionInfo, context: { sourceVersionId: string }): Promise<void> | void
}

export interface HistoryRestorePreparation {
  activate(version: HistoryVersionInfo, context: { sourceVersionId: string }): Promise<void> | void
  cancel(cause: unknown): Promise<void> | void
}

export class HistoryPreparationReleaseError extends Error {
  constructor(readonly restoreError: unknown, readonly releaseError: unknown) {
    super('History restore failed and the prepared workspace could not resume', { cause: restoreError })
    this.name = 'HistoryPreparationReleaseError'
  }
}

export interface CaptureHistoryCommandOptions extends Omit<HistoryCaptureOptions<never>, 'snapshot' | 'contentType' | 'signal'> {
  signal?: AbortSignal
}

export interface RestoreHistoryCommandOptions extends Omit<HistoryRestoreOptions, 'signal'> {
  signal?: AbortSignal
}

export class HistoryCommandBusyError extends Error {
  constructor() { super('A history command is already running'); this.name = 'HistoryCommandBusyError' }
}

export class HistoryActivationError extends Error {
  constructor(readonly version: HistoryVersionInfo, cause: unknown) {
    super(`Durable restore ${version.id} was created but the live workspace could not activate it`, { cause })
    this.name = 'HistoryActivationError'
  }
}

type Listener = (state: Readonly<HistoryCommandState>) => void

/** Serialized history workflow controller for hosts and Univer commands. */
export class HistoryCommandController<TSnapshot> {
  private state: HistoryCommandState = { jobId: null, operation: null, status: 'idle' }
  private active: AbortController | null = null
  private sequence = 0
  private readonly listeners = new Set<Listener>()

  constructor(readonly manager: HistoryManager<TSnapshot>, private readonly workspace: HistoryWorkspace<TSnapshot>) {}

  onState(listener: Listener): () => void { this.listeners.add(listener); return () => this.listeners.delete(listener) }
  getState(): HistoryCommandState { return { ...this.state } }
  cancel(): boolean { if (!this.active) return false; this.active.abort(); return true }

  list(options: HistoryListOptions = {}): Promise<HistoryListPage> {
    return this.run('list', options.signal, (signal) => this.manager.listVersions({ ...options, signal }))
  }

  preview(versionId: string, signal?: AbortSignal): Promise<HistoryPreview<TSnapshot>> {
    return this.run('preview', signal, async (jobSignal) => {
      const preview = await this.manager.loadPreview(versionId, jobSignal)
      if (jobSignal.aborted) throw abortError()
      await this.workspace.showIsolatedPreview(preview, { signal: jobSignal })
      if (jobSignal.aborted) throw abortError()
      return preview
    })
  }

  capture(options: CaptureHistoryCommandOptions): Promise<HistoryVersionInfo> {
    return this.run('capture', options.signal, async (signal) => {
      const captured = await this.workspace.captureSnapshot({ signal })
      if (signal.aborted) throw abortError()
      return this.manager.capture({ ...options, snapshot: captured.snapshot, contentType: captured.contentType, signal })
    })
  }

  restore(options: RestoreHistoryCommandOptions): Promise<HistoryVersionInfo> {
    return this.run('restore', options.signal, async (signal) => {
      const preparation = await this.workspace.prepareRestore?.({
        sourceVersionId: options.sourceVersionId,
        expectedHeadVersionId: options.expectedHeadVersionId,
        signal,
      })
      let version: HistoryVersionInfo
      try {
        version = await this.manager.restore({ ...options, signal })
      } catch (error) {
        if (preparation) {
          try { await preparation.cancel(error) } catch (releaseError) {
            throw new HistoryPreparationReleaseError(error, releaseError)
          }
        }
        throw error
      }
      try {
        if (preparation) await preparation.activate(version, { sourceVersionId: options.sourceVersionId })
        else await this.workspace.activateRestoredVersion(version, { sourceVersionId: options.sourceVersionId })
      } catch (error) {
        throw new HistoryActivationError(version, error)
      }
      return version
    })
  }

  private async run<T>(operation: HistoryCommandOperation, externalSignal: AbortSignal | undefined, action: (signal: AbortSignal) => Promise<T>): Promise<T> {
    // A cancelled host request may take time to settle. Let the replacement
    // begin immediately while preventing the stale job from overwriting state.
    if (this.active && !this.active.signal.aborted) throw new HistoryCommandBusyError()
    const controller = new AbortController()
    const abort = () => controller.abort()
    externalSignal?.addEventListener('abort', abort, { once: true })
    if (externalSignal?.aborted) controller.abort()
    this.active = controller
    const jobId = `history-${++this.sequence}`
    this.setState({ jobId, operation, status: 'running' })
    try {
      if (controller.signal.aborted) throw abortError()
      const result = await action(controller.signal)
      if (controller.signal.aborted && operation !== 'restore') throw abortError()
      if (this.active === controller) this.setState({ jobId, operation, status: 'succeeded' })
      return result
    } catch (error) {
      if (this.active === controller) {
        if (controller.signal.aborted && !(error instanceof HistoryActivationError)) {
          this.setState({ jobId, operation, status: 'cancelled' })
        } else {
          this.setState({ jobId, operation, status: 'failed', error: error instanceof Error ? error.message : String(error) })
        }
      }
      throw error
    } finally {
      externalSignal?.removeEventListener('abort', abort)
      if (this.active === controller) this.active = null
    }
  }

  private setState(state: HistoryCommandState): void {
    this.state = state
    for (const listener of this.listeners) listener({ ...state })
  }
}

function abortError(): Error {
  return new DOMException('History command cancelled', 'AbortError')
}

/** Convenience author guard for browser command adapters. */
export function isHistoryAuthor(value: unknown): value is HistoryAuthor {
  if (!value || typeof value !== 'object') return false
  const author = value as Partial<HistoryAuthor>
  return typeof author.id === 'string' && author.id.length > 0
    && (author.kind === 'user' || author.kind === 'agent' || author.kind === 'system')
    && (author.displayName === undefined || typeof author.displayName === 'string')
}
