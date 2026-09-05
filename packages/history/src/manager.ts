import type {
  HistoryCaptureOptions,
  HistoryCreateRequest,
  HistoryEvent,
  HistoryHost,
  HistoryListOptions,
  HistoryListPage,
  HistoryPreview,
  HistoryRestoreOptions,
  HistoryRetention,
  HistoryVersionInfo,
} from './types'
import {
  historyVersionMatches,
  validateHistoryAuthor,
  validateHistoryListOptions,
  validateHistoryListPage,
  validateHistoryVersionInfo,
  validateRequestedRetention,
} from './validation'

type HistoryListener = (event: Readonly<HistoryEvent>) => void

export class HistoryProtocolError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'HistoryProtocolError'
  }
}

function assertId(value: string, label: string): void {
  if (typeof value !== 'string' || value.trim().length === 0 || value.length > 512) throw new TypeError(`${label} must contain 1-512 characters`)
}

function assertDescription(value: string | undefined): void {
  if (value !== undefined && (typeof value !== 'string' || value.length > 2048)) throw new TypeError('description must not exceed 2048 characters')
}

function clone<T>(value: T, label: string): T {
  try {
    return structuredClone(value)
  } catch {
    throw new TypeError(`${label} must be structured-cloneable`)
  }
}

function failure(prefix: string, issues: { path: string; message: string }[]): HistoryProtocolError {
  return new HistoryProtocolError(`${prefix}: ${issues.map((entry) => `${entry.path} ${entry.message}`).join('; ')}`)
}

/** Host-backed durable version lifecycle. This manager never mutates the live editor document. */
export class HistoryManager<TSnapshot> {
  private readonly listeners = new Set<HistoryListener>()

  constructor(private readonly host: HistoryHost<TSnapshot>, readonly artifactId: string) {
    assertId(artifactId, 'artifactId')
  }

  onEvent(listener: HistoryListener): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  async listVersions(options: HistoryListOptions = {}): Promise<HistoryListPage> {
    const checked = validateHistoryListOptions(options)
    if (!checked.ok) throw new TypeError(checked.issues.map((entry) => `${entry.path} ${entry.message}`).join('; '))
    const { signal, ...data } = checked.value
    const request = { ...clone({ artifactId: this.artifactId, ...data }, 'history list request'), ...(signal === undefined ? {} : { signal }) }
    const response = await this.host.listVersions(request)
    const result = validateHistoryListPage(response, this.artifactId)
    if (!result.ok) throw failure('Invalid history host list response', result.issues)
    if (result.value.versions.some((version) => !historyVersionMatches(version, checked.value.filter))) {
      throw new HistoryProtocolError('Invalid history host list response: host did not apply the requested filter before pagination')
    }
    return clone(result.value, 'history list response')
  }

  async loadPreview(versionId: string, signal?: AbortSignal): Promise<HistoryPreview<TSnapshot>> {
    assertId(versionId, 'versionId')
    const loaded = await this.host.loadVersion({ artifactId: this.artifactId, versionId, signal })
    if (!loaded || typeof loaded !== 'object') throw new HistoryProtocolError('Invalid history host load response: response must be an object')
    const result = validateHistoryVersionInfo(loaded.version)
    if (!result.ok) throw failure('Invalid history host load response', result.issues)
    if (result.value.artifactId !== this.artifactId || result.value.id !== versionId) throw new HistoryProtocolError('Invalid history host load response: version identity does not match the request')
    return { mode: 'isolated', version: clone(result.value, 'history version'), snapshot: clone(loaded.snapshot, 'history snapshot') }
  }

  async capture(options: HistoryCaptureOptions<TSnapshot>): Promise<HistoryVersionInfo> {
    const reason = options.reason ?? 'save'
    if (!['save', 'agent-delivery', 'import'].includes(reason)) throw new TypeError('capture reason must be save, agent-delivery, or import')
    const request = this.createRequest({ ...options, reason })
    const version = await this.host.createVersion(request)
    const checked = this.validateCreated(version, request)
    this.emit({ name: 'captured', version: checked })
    return clone(checked, 'history version')
  }

  async restore(options: HistoryRestoreOptions): Promise<HistoryVersionInfo> {
    assertId(options.sourceVersionId, 'sourceVersionId')
    const source = await this.loadPreview(options.sourceVersionId, options.signal)
    if (options.signal?.aborted) throw new DOMException('History restore cancelled before durable creation', 'AbortError')
    const request = this.createRequest({
      snapshot: source.snapshot,
      author: options.author,
      reason: 'restore',
      sourceVersionId: options.sourceVersionId,
      expectedHeadVersionId: options.expectedHeadVersionId,
      description: options.description,
      retention: options.retention,
      contentType: source.version.contentType,
      signal: options.signal,
    })
    const version = await this.host.createVersion(request)
    const checked = this.validateCreated(version, request)
    if (checked.id === options.sourceVersionId) throw new HistoryProtocolError('Invalid history host create response: restore must create a new immutable version')
    this.emit({ name: 'restored', version: checked, sourceVersionId: options.sourceVersionId })
    return clone(checked, 'history version')
  }

  private createRequest(options: Omit<HistoryCreateRequest<TSnapshot>, 'artifactId'>): HistoryCreateRequest<TSnapshot> {
    const author = validateHistoryAuthor(options.author)
    if (!author.ok) throw new TypeError(author.issues.map((entry) => `${entry.path} ${entry.message}`).join('; '))
    assertId(options.contentType, 'contentType')
    if (options.expectedHeadVersionId !== undefined) assertId(options.expectedHeadVersionId, 'expectedHeadVersionId')
    if (options.sourceVersionId !== undefined) assertId(options.sourceVersionId, 'sourceVersionId')
    assertDescription(options.description)
    let retention: Partial<HistoryRetention> | undefined
    if (options.retention !== undefined) {
      const checked = validateRequestedRetention(options.retention)
      if (!checked.ok) throw new TypeError(checked.issues.map((entry) => `${entry.path} ${entry.message}`).join('; '))
      retention = checked.value
    }
    const request = clone({
      artifactId: this.artifactId,
      author: author.value,
      reason: options.reason,
      contentType: options.contentType,
      snapshot: options.snapshot,
      ...(options.expectedHeadVersionId === undefined ? {} : { expectedHeadVersionId: options.expectedHeadVersionId }),
      ...(options.sourceVersionId === undefined ? {} : { sourceVersionId: options.sourceVersionId }),
      ...(options.description === undefined ? {} : { description: options.description }),
      ...(retention === undefined ? {} : { retention }),
    }, 'history create request')
    return { ...request, ...(options.signal === undefined ? {} : { signal: options.signal }) }
  }

  private validateCreated(input: unknown, request: HistoryCreateRequest<TSnapshot>): HistoryVersionInfo {
    const result = validateHistoryVersionInfo(input)
    if (!result.ok) throw failure('Invalid history host create response', result.issues)
    const version = result.value
    if (version.artifactId !== this.artifactId || version.reason !== request.reason || version.sourceVersionId !== request.sourceVersionId
      || version.author.id !== request.author.id || version.author.kind !== request.author.kind || version.contentType !== request.contentType) {
      throw new HistoryProtocolError('Invalid history host create response: artifact, author, content type, reason, or restore lineage does not match the request')
    }
    return version
  }

  private emit(event: HistoryEvent): void {
    for (const listener of this.listeners) listener(clone(event, 'history event'))
  }
}
