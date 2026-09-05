import type { PrintSnapshot } from './types'
import { validatePrintLayout, validatePrintRender } from './validation'

export type PrintPageSelection =
  | { kind: 'all' }
  | { kind: 'pages'; pages: ReadonlyArray<number> }
  | { kind: 'range'; startPage: number; endPage: number }

export interface PrintPreviewPage<TPayload = unknown> {
  /** One-based page number in the complete rendered document. */
  number: number
  widthPoints: number
  heightPoints: number
  /** Host-defined, structured-cloneable render handle or payload. */
  payload: TPayload
}

export interface PrintPreviewOutput<TPayload = unknown> {
  totalPages: number
  pages: ReadonlyArray<PrintPreviewPage<TPayload>>
}

export interface PrintPreviewRequest {
  sessionId: string
  snapshot: Readonly<PrintSnapshot>
  selection: Readonly<PrintPageSelection>
}

export interface PrintPreviewHost<TPayload = unknown> {
  renderPreview(request: Readonly<PrintPreviewRequest>, context: Readonly<{ signal: AbortSignal }>): Promise<PrintPreviewOutput<TPayload>> | PrintPreviewOutput<TPayload>
}

export type PrintPreviewState = 'rendering' | 'ready' | 'canceled' | 'timed-out' | 'failed'

export interface PrintPreviewStateSnapshot {
  sessionId: string
  state: PrintPreviewState
  selection: PrintPageSelection
}

export interface PrintPreviewDocument<TPayload = unknown> extends PrintPreviewOutput<TPayload> {
  sessionId: string
  snapshot: PrintSnapshot
  selection: PrintPageSelection
}

export interface PrintPreviewSession<TPayload = unknown> {
  readonly id: string
  readonly result: Promise<PrintPreviewDocument<TPayload>>
  cancel(reason?: unknown): boolean
  getState(): PrintPreviewStateSnapshot
}

export type PrintPreviewEvent<TPayload = unknown> =
  | { type: 'started'; session: PrintPreviewStateSnapshot }
  | { type: 'ready'; session: PrintPreviewStateSnapshot; document: PrintPreviewDocument<TPayload> }
  | { type: 'canceled' | 'timed-out' | 'failed'; session: PrintPreviewStateSnapshot; error: PrintPreviewError }

export type PrintPreviewErrorCode = 'ABORTED' | 'DUPLICATE_SESSION' | 'HOST_FAILED' | 'INVALID_OUTPUT' | 'INVALID_REQUEST' | 'TIMEOUT'

export class PrintPreviewError extends Error {
  constructor(public readonly code: PrintPreviewErrorCode, message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'PrintPreviewError'
  }
}

export interface PrintPreviewStartOptions {
  selection?: PrintPageSelection
  signal?: AbortSignal
  timeoutMs?: number
}

export interface PrintPreviewManagerOptions {
  timeoutMs?: number
}

function clone<T>(value: T, label: string, code: 'INVALID_REQUEST' | 'INVALID_OUTPUT' = 'INVALID_REQUEST'): T {
  try {
    return structuredClone(value)
  } catch {
    throw new PrintPreviewError(code, `${label} must be structured-cloneable`)
  }
}

function validateTimeout(value: number): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > 3_600_000) throw new PrintPreviewError('INVALID_REQUEST', 'timeoutMs must be an integer from 1 through 3600000')
  return value
}

function validateSessionId(value: string): void {
  if (typeof value !== 'string' || value.length === 0 || value.length > 256 || !/^[A-Za-z0-9][A-Za-z0-9._:/-]*$/.test(value)) throw new PrintPreviewError('INVALID_REQUEST', 'sessionId must be a stable identifier containing 1-256 characters')
}

function canonicalSelection(input: PrintPageSelection | undefined): PrintPageSelection {
  if (input === undefined) return { kind: 'all' }
  if (!input || typeof input !== 'object') throw new PrintPreviewError('INVALID_REQUEST', 'page selection must be an object')
  if (input.kind === 'all') return { kind: 'all' }
  if (input.kind === 'range') {
    if (!Number.isSafeInteger(input.startPage) || !Number.isSafeInteger(input.endPage) || input.startPage < 1 || input.endPage < input.startPage || input.endPage > 100_000) throw new PrintPreviewError('INVALID_REQUEST', 'page range must be ordered, one-based, and no greater than 100000')
    return { kind: 'range', startPage: input.startPage, endPage: input.endPage }
  }
  if (input.kind !== 'pages' || !Array.isArray(input.pages) || input.pages.length === 0 || input.pages.length > 100_000) throw new PrintPreviewError('INVALID_REQUEST', 'pages selection must contain 1-100000 page numbers')
  const pages = [...input.pages]
  if (pages.some((page) => !Number.isSafeInteger(page) || page < 1 || page > 100_000) || new Set(pages).size !== pages.length) throw new PrintPreviewError('INVALID_REQUEST', 'page numbers must be unique one-based integers no greater than 100000')
  pages.sort((left, right) => left - right)
  return { kind: 'pages', pages }
}

function selectedPages(selection: PrintPageSelection, totalPages: number): number[] {
  if (selection.kind === 'all') return Array.from({ length: totalPages }, (_, index) => index + 1)
  if (selection.kind === 'range') {
    if (selection.endPage > totalPages) throw new PrintPreviewError('INVALID_OUTPUT', 'selected page range exceeds totalPages')
    return Array.from({ length: selection.endPage - selection.startPage + 1 }, (_, index) => selection.startPage + index)
  }
  if (selection.pages.some((page) => page > totalPages)) throw new PrintPreviewError('INVALID_OUTPUT', 'selected page number exceeds totalPages')
  return [...selection.pages]
}

function validateOutput<TPayload>(input: PrintPreviewOutput<TPayload>, selection: PrintPageSelection): PrintPreviewOutput<TPayload> {
  if (!input || typeof input !== 'object' || !Number.isSafeInteger(input.totalPages) || input.totalPages < 0 || input.totalPages > 100_000 || !Array.isArray(input.pages)) throw new PrintPreviewError('INVALID_OUTPUT', 'preview output requires bounded totalPages and a pages array')
  const expected = selectedPages(selection, input.totalPages)
  if (input.pages.length !== expected.length) throw new PrintPreviewError('INVALID_OUTPUT', 'host must return exactly the selected preview pages')
  const seen = new Set<number>()
  input.pages.forEach((page, index) => {
    if (!page || typeof page !== 'object' || !Number.isSafeInteger(page.number) || page.number < 1 || page.number > input.totalPages || seen.has(page.number)) throw new PrintPreviewError('INVALID_OUTPUT', `pages[${index}] has an invalid or duplicate page number`)
    if (!Number.isFinite(page.widthPoints) || page.widthPoints <= 0 || page.widthPoints > 20_000 || !Number.isFinite(page.heightPoints) || page.heightPoints <= 0 || page.heightPoints > 20_000) throw new PrintPreviewError('INVALID_OUTPUT', `pages[${index}] has invalid point dimensions`)
    seen.add(page.number)
  })
  if (expected.some((page) => !seen.has(page))) throw new PrintPreviewError('INVALID_OUTPUT', 'host returned pages outside the requested selection')
  const cloned = clone(input, 'preview output', 'INVALID_OUTPUT')
  return { totalPages: cloned.totalPages, pages: [...cloned.pages].sort((left, right) => left.number - right.number) }
}

/** Host-facing preview-session lifecycle. This class validates pages; it does not render or invoke browser print APIs. */
export class PrintPreviewManager<TPayload = unknown> {
  private readonly sessions = new Map<string, { state: PrintPreviewStateSnapshot; controller: AbortController }>()
  private readonly listeners = new Set<(event: Readonly<PrintPreviewEvent<TPayload>>) => void>()
  private readonly defaultTimeoutMs: number

  constructor(private readonly host: PrintPreviewHost<TPayload>, options: PrintPreviewManagerOptions = {}) {
    if (!host || typeof host.renderPreview !== 'function') throw new PrintPreviewError('INVALID_REQUEST', 'preview host must expose renderPreview()')
    this.defaultTimeoutMs = validateTimeout(options.timeoutMs ?? 30_000)
  }

  onEvent(listener: (event: Readonly<PrintPreviewEvent<TPayload>>) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  start(sessionId: string, snapshot: PrintSnapshot, options: PrintPreviewStartOptions = {}): PrintPreviewSession<TPayload> {
    validateSessionId(sessionId)
    if (this.sessions.has(sessionId)) throw new PrintPreviewError('DUPLICATE_SESSION', `preview session ${sessionId} already exists`)
    if (!snapshot || typeof snapshot !== 'object' || !snapshot.layout || !snapshot.render) throw new PrintPreviewError('INVALID_REQUEST', 'snapshot must contain print layout and render configuration')
    const layout = validatePrintLayout(snapshot.layout)
    const render = validatePrintRender(snapshot.render)
    if (!layout.ok || !render.ok || typeof snapshot.dialogOpen !== 'boolean') throw new PrintPreviewError('INVALID_REQUEST', 'snapshot contains invalid print layout, render, or dialog state')
    const canonicalSnapshot = clone({ layout: layout.value, render: render.value, dialogOpen: snapshot.dialogOpen }, 'print snapshot')
    const selection = canonicalSelection(options.selection)
    const timeoutMs = validateTimeout(options.timeoutMs ?? this.defaultTimeoutMs)
    const controller = new AbortController()
    const record = { state: { sessionId, state: 'rendering' as PrintPreviewState, selection }, controller }
    this.sessions.set(sessionId, record)
    this.emit({ type: 'started', session: record.state })
    const result = this.run(record, canonicalSnapshot, selection, options.signal, timeoutMs)
    return {
      id: sessionId,
      result,
      cancel: (reason?: unknown) => this.cancel(sessionId, reason),
      getState: () => clone(record.state, 'preview state'),
    }
  }

  cancel(sessionId: string, reason?: unknown): boolean {
    const record = this.sessions.get(sessionId)
    if (!record || record.state.state !== 'rendering') return false
    record.controller.abort(reason)
    return true
  }

  private async run(record: { state: PrintPreviewStateSnapshot; controller: AbortController }, snapshot: PrintSnapshot, selection: PrintPageSelection, signal: AbortSignal | undefined, timeoutMs: number): Promise<PrintPreviewDocument<TPayload>> {
    const externalAbort = () => record.controller.abort(signal?.reason)
    signal?.addEventListener('abort', externalAbort, { once: true })
    if (signal?.aborted) externalAbort()
    let timedOut = false
    const timer = setTimeout(() => {
      timedOut = true
      record.controller.abort()
    }, timeoutMs)
    try {
      if (record.controller.signal.aborted) throw new PrintPreviewError('ABORTED', 'preview was canceled before rendering')
      const aborted = new Promise<never>((_, reject) => record.controller.signal.addEventListener('abort', () => reject(new PrintPreviewError(timedOut ? 'TIMEOUT' : 'ABORTED', timedOut ? `preview exceeded ${timeoutMs}ms` : 'preview was canceled')), { once: true }))
      let output: PrintPreviewOutput<TPayload>
      try {
        const hostSnapshot = clone(snapshot, 'preview host snapshot')
        const hostSelection = clone(selection, 'preview host selection')
        output = await Promise.race([
          Promise.resolve(this.host.renderPreview(Object.freeze({ sessionId: record.state.sessionId, snapshot: hostSnapshot, selection: hostSelection }), { signal: record.controller.signal })),
          aborted,
        ])
      } catch (cause) {
        if (cause instanceof PrintPreviewError) throw cause
        if (record.controller.signal.aborted) throw new PrintPreviewError(timedOut ? 'TIMEOUT' : 'ABORTED', timedOut ? `preview exceeded ${timeoutMs}ms` : 'preview was canceled', { cause })
        throw new PrintPreviewError('HOST_FAILED', 'preview host failed', { cause })
      }
      const checked = validateOutput(output, selection)
      const document = clone({ sessionId: record.state.sessionId, snapshot, selection, ...checked }, 'preview document', 'INVALID_OUTPUT')
      record.state = { ...record.state, state: 'ready' }
      this.emit({ type: 'ready', session: record.state, document })
      return document
    } catch (cause) {
      const error = cause instanceof PrintPreviewError ? cause : new PrintPreviewError('HOST_FAILED', 'preview host failed', { cause })
      const state: PrintPreviewState = error.code === 'ABORTED' ? 'canceled' : error.code === 'TIMEOUT' ? 'timed-out' : 'failed'
      record.state = { ...record.state, state }
      this.emit({ type: state === 'canceled' ? 'canceled' : state === 'timed-out' ? 'timed-out' : 'failed', session: record.state, error })
      throw error
    } finally {
      clearTimeout(timer)
      signal?.removeEventListener('abort', externalAbort)
      this.sessions.delete(record.state.sessionId)
    }
  }

  private emit(event: PrintPreviewEvent<TPayload>): void {
    for (const listener of this.listeners) {
      try {
        if ('error' in event) listener({ type: event.type, session: clone(event.session, 'preview event'), error: new PrintPreviewError(event.error.code, event.error.message) })
        else listener(clone(event, 'preview event', 'INVALID_OUTPUT'))
      } catch {
        // Observers cannot change preview lifecycle semantics.
      }
    }
  }
}
