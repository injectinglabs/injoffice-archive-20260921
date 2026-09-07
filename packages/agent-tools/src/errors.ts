import type { AgentErrorCode, AgentErrorData, AgentIssue, JsonObject } from './types'

export class AgentToolsError extends Error {
  readonly name = 'AgentToolsError'

  constructor(
    readonly code: AgentErrorCode,
    message: string,
    readonly retryable = false,
    readonly issues?: readonly AgentIssue[],
    readonly details?: JsonObject,
    options?: ErrorOptions,
  ) {
    super(message, options)
  }

  toJSON(): AgentErrorData {
    return {
      code: this.code,
      message: this.message,
      retryable: this.retryable,
      ...(this.issues ? { issues: this.issues } : {}),
      ...(this.details ? { details: this.details } : {}),
    }
  }
}

export function asAgentToolsError(error: unknown): AgentToolsError {
  if (error instanceof AgentToolsError) return error
  if (isAbortError(error)) return new AgentToolsError('ABORTED', 'Agent operation was aborted.', true, undefined, undefined, { cause: error })
  return new AgentToolsError('ADAPTER_ERROR', error instanceof Error ? error.message : 'Artifact adapter failed.', true, undefined, undefined, { cause: error })
}

function isAbortError(error: unknown): boolean {
  return typeof error === 'object' && error !== null && 'name' in error && error.name === 'AbortError'
}
