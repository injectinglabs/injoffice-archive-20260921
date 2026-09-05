export type VersionAuthor = 'agent' | 'user'

/** @deprecated Use HistoryVersionInfo for the durable history lifecycle. */
export interface VersionInfo {
  id: string
  author: VersionAuthor
  ts: number
  size: number
}

/** @deprecated Use HistoryListPage for the durable history lifecycle. */
export interface VersionListResponse {
  versions: VersionInfo[]
}

export type HistoryAuthorKind = 'user' | 'agent' | 'system'
export type HistoryChangeReason = 'save' | 'agent-delivery' | 'import' | 'restore'

export interface HistoryAuthor {
  /** Stable host identity. This is not a display name. */
  id: string
  kind: HistoryAuthorKind
  displayName?: string
}

/** Effective server-owned retention state for one immutable version. */
export interface HistoryRetention {
  policyId: string
  /** Unix milliseconds, or null when the policy has no scheduled expiry. */
  expiresAt: number | null
  legalHold: boolean
}

export interface HistoryVersionInfo {
  /** Opaque stable identifier used for subsequent load and restore requests. */
  id: string
  artifactId: string
  /** Monotonically increasing within an artifact. */
  sequence: number
  createdAt: number
  size: number
  contentType: string
  author: HistoryAuthor
  reason: HistoryChangeReason
  /** Present only when this version restored an earlier immutable version. */
  sourceVersionId?: string
  description?: string
  retention: HistoryRetention
}

export interface HistoryVersionFilter {
  authorIds?: string[]
  authorKinds?: HistoryAuthorKind[]
  reasons?: HistoryChangeReason[]
  createdAfter?: number
  createdBefore?: number
  sourceVersionId?: string
}

export interface HistoryListOptions {
  cursor?: string
  limit?: number
  filter?: HistoryVersionFilter
  signal?: AbortSignal
}

export interface HistoryListRequest extends HistoryListOptions {
  artifactId: string
}

export interface HistoryListPage {
  /** Newest first. Filtering must happen before host pagination. */
  versions: HistoryVersionInfo[]
  nextCursor?: string
}

export interface HistoryLoadRequest {
  artifactId: string
  versionId: string
  signal?: AbortSignal
}

export interface HistoryLoadedVersion<TSnapshot> {
  version: HistoryVersionInfo
  snapshot: TSnapshot
}

export interface HistoryCreateRequest<TSnapshot> {
  artifactId: string
  /** Enables the host to reject a stale capture or restore. */
  expectedHeadVersionId?: string
  author: HistoryAuthor
  reason: HistoryChangeReason
  sourceVersionId?: string
  description?: string
  /** Requested policy; the returned version contains the effective policy. */
  retention?: Partial<HistoryRetention>
  contentType: string
  snapshot: TSnapshot
  signal?: AbortSignal
}

/**
 * Persistence boundary supplied by an application server. Implementations must
 * store immutable snapshots and apply list filters before pagination.
 */
export interface HistoryHost<TSnapshot> {
  listVersions(request: Readonly<HistoryListRequest>): Promise<HistoryListPage>
  loadVersion(request: Readonly<HistoryLoadRequest>): Promise<HistoryLoadedVersion<TSnapshot>>
  createVersion(request: Readonly<HistoryCreateRequest<TSnapshot>>): Promise<HistoryVersionInfo>
}

export interface HistoryCaptureOptions<TSnapshot> {
  snapshot: TSnapshot
  author: HistoryAuthor
  reason?: Exclude<HistoryChangeReason, 'restore'>
  expectedHeadVersionId?: string
  description?: string
  retention?: Partial<HistoryRetention>
  contentType: string
  signal?: AbortSignal
}

export interface HistoryRestoreOptions {
  sourceVersionId: string
  author: HistoryAuthor
  expectedHeadVersionId?: string
  description?: string
  retention?: Partial<HistoryRetention>
  signal?: AbortSignal
}

/** A detached snapshot that must not be mounted as the live editor document. */
export interface HistoryPreview<TSnapshot> extends HistoryLoadedVersion<TSnapshot> {
  mode: 'isolated'
}

export type HistoryEventName = 'captured' | 'restored'

export interface HistoryEvent {
  name: HistoryEventName
  version: HistoryVersionInfo
  sourceVersionId?: string
}

export interface HistoryIssue {
  path: string
  code: 'DUPLICATE' | 'INVALID_ENUM' | 'INVALID_NUMBER' | 'INVALID_ORDER' | 'INVALID_VALUE' | 'REQUIRED'
  message: string
}

export type HistoryResult<T> = { ok: true; value: T } | { ok: false; issues: HistoryIssue[] }
