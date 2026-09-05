import type { HistoryRestorePreparation } from './commands'
import type { HistoryVersionInfo } from './types'

export type HistoryCollaborationPhase = 'active' | 'quiesced' | 'recovering'

export interface HistoryCollaborationState {
  room: string
  artifactVersionId: string
  logEpoch: string
  headSequence: number
  savedSequence: number
  pendingOutbound: number
  phase: HistoryCollaborationPhase
}

export interface HistoryCollaborationBoundary {
  authorizeRestore(context: { room: string; sourceVersionId: string; expectedHeadVersionId?: string; signal: AbortSignal }): Promise<boolean> | boolean
  /** Stop accepting local mutations and flush durable outbound work. */
  quiesce(context: { room: string; signal: AbortSignal }): Promise<HistoryCollaborationState>
  /** Atomically load the durable version and reset the old operation epoch. */
  reloadAndReset(context: { room: string; version: HistoryVersionInfo; sourceVersionId: string }): Promise<void>
  /** Join the new epoch and return the server-authoritative state. */
  rejoin(context: { room: string; versionId: string }): Promise<HistoryCollaborationState>
  /** Resume the unchanged pre-restore session when durable creation fails. */
  resume(context: { room: string; cause: unknown }): Promise<HistoryCollaborationState>
  onRecoveryRequired?(context: { room: string; version: HistoryVersionInfo; cause: unknown }): void
}

export type HistoryCollaborationConflictCode =
  | 'busy'
  | 'permission'
  | 'identity'
  | 'pending-outbound'
  | 'unsaved-head'
  | 'phase'
  | 'invalid-sequence'

export class HistoryCollaborationConflictError extends Error {
  constructor(readonly code: HistoryCollaborationConflictCode, message: string) {
    super(message)
    this.name = 'HistoryCollaborationConflictError'
  }
}

/** Coordinates restore with a host's collaboration room without assuming a
 * particular transport or server. Restore is refused until the current epoch
 * is quiescent, fully acknowledged, and durably represented by its head. */
export class HistoryCollaborationCoordinator {
  private prepared = false

  constructor(readonly room: string, private readonly boundary: HistoryCollaborationBoundary) {
    if (!room.trim()) throw new TypeError('collaboration room must not be empty')
  }

  async prepare(context: {
    sourceVersionId: string
    expectedHeadVersionId?: string
    signal: AbortSignal
  }): Promise<HistoryRestorePreparation> {
    if (this.prepared) throw new HistoryCollaborationConflictError('busy', 'A collaborative restore is already prepared')
    this.prepared = true
    let quiesced = false
    let before: HistoryCollaborationState
    try {
      const allowed = await this.boundary.authorizeRestore({ room: this.room, ...context })
      if (!allowed) throw new HistoryCollaborationConflictError('permission', 'The current user cannot restore this collaboration room')
      before = await this.boundary.quiesce({ room: this.room, signal: context.signal })
      quiesced = true
      validateState(before, this.room)
      if (before.phase !== 'quiesced') throw new HistoryCollaborationConflictError('phase', 'Collaboration did not enter the quiesced phase')
      if (before.pendingOutbound !== 0) throw new HistoryCollaborationConflictError('pending-outbound', 'Outbound collaboration edits remain pending')
      if (before.savedSequence !== before.headSequence) throw new HistoryCollaborationConflictError('unsaved-head', 'The collaboration head is not durably saved')
      if (context.expectedHeadVersionId !== undefined && before.artifactVersionId !== context.expectedHeadVersionId) {
        throw new HistoryCollaborationConflictError('identity', 'The collaboration artifact version does not match the expected history head')
      }
    } catch (error) {
      if (quiesced) {
        try { await this.boundary.resume({ room: this.room, cause: error }) } catch { /* original refusal remains authoritative */ }
      }
      this.prepared = false
      throw error
    }

    let consumed = false
    const consume = () => {
      if (consumed) throw new HistoryCollaborationConflictError('phase', 'The restore preparation has already been consumed')
      consumed = true
    }
    return {
      activate: async (version, activation) => {
        consume()
        try {
          await this.boundary.reloadAndReset({ room: this.room, version, sourceVersionId: activation.sourceVersionId })
          const joined = await this.boundary.rejoin({ room: this.room, versionId: version.id })
          validateState(joined, this.room)
          if (joined.phase !== 'active' || joined.artifactVersionId !== version.id || joined.pendingOutbound !== 0) {
            throw new HistoryCollaborationConflictError('identity', 'Rejoined collaboration state does not represent the restored version')
          }
        } catch (error) {
          this.boundary.onRecoveryRequired?.({ room: this.room, version, cause: error })
          throw error
        } finally {
          this.prepared = false
        }
      },
      cancel: async (cause) => {
        consume()
        try {
          const resumed = await this.boundary.resume({ room: this.room, cause })
          validateState(resumed, this.room)
          if (resumed.phase !== 'active' || resumed.artifactVersionId !== before.artifactVersionId) {
            throw new HistoryCollaborationConflictError('identity', 'Resumed collaboration state no longer represents the original version')
          }
        } finally {
          this.prepared = false
        }
      },
    }
  }
}

function validateState(state: HistoryCollaborationState, room: string): void {
  if (!state || typeof state !== 'object' || state.room !== room || !state.artifactVersionId?.trim() || !state.logEpoch?.trim()) {
    throw new HistoryCollaborationConflictError('identity', 'Collaboration state has invalid room, artifact, or epoch identity')
  }
  if (![state.headSequence, state.savedSequence, state.pendingOutbound].every((value) => Number.isSafeInteger(value) && value >= 0)
    || state.savedSequence > state.headSequence) {
    throw new HistoryCollaborationConflictError('invalid-sequence', 'Collaboration state has invalid sequence or pending counts')
  }
  if (!['active', 'quiesced', 'recovering'].includes(state.phase)) {
    throw new HistoryCollaborationConflictError('phase', 'Collaboration state has an invalid phase')
  }
}
