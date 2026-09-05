import type {
  HistoryAuthor,
  HistoryAuthorKind,
  HistoryChangeReason,
  HistoryIssue,
  HistoryListOptions,
  HistoryListPage,
  HistoryResult,
  HistoryRetention,
  HistoryVersionFilter,
  HistoryVersionInfo,
} from './types'

const AUTHOR_KINDS = new Set<HistoryAuthorKind>(['user', 'agent', 'system'])
const REASONS = new Set<HistoryChangeReason>(['save', 'agent-delivery', 'import', 'restore'])
const MAX_ID = 512
const MAX_TEXT = 2048

function issue(issues: HistoryIssue[], path: string, code: HistoryIssue['code'], message: string): void {
  issues.push({ path, code, message })
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function text(value: unknown, path: string, issues: HistoryIssue[], maximum = MAX_ID): value is string {
  if (typeof value !== 'string' || value.trim().length === 0 || value.length > maximum) {
    issue(issues, path, 'REQUIRED', `must contain 1-${maximum} characters`)
    return false
  }
  return true
}

function timestamp(value: unknown, path: string, issues: HistoryIssue[]): value is number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    issue(issues, path, 'INVALID_NUMBER', 'must be a non-negative unix-millisecond integer')
    return false
  }
  return true
}

function validateAuthorInto(value: unknown, path: string, issues: HistoryIssue[]): value is HistoryAuthor {
  if (!record(value)) {
    issue(issues, path, 'REQUIRED', 'author must be an object')
    return false
  }
  text(value.id, `${path}/id`, issues)
  if (!AUTHOR_KINDS.has(value.kind as HistoryAuthorKind)) issue(issues, `${path}/kind`, 'INVALID_ENUM', 'unsupported author kind')
  if (value.displayName !== undefined && (typeof value.displayName !== 'string' || value.displayName.length > 256)) {
    issue(issues, `${path}/displayName`, 'INVALID_VALUE', 'display name must not exceed 256 characters')
  }
  return true
}

function validateRetentionInto(value: unknown, path: string, issues: HistoryIssue[], createdAt?: number): value is HistoryRetention {
  if (!record(value)) {
    issue(issues, path, 'REQUIRED', 'effective retention metadata is required')
    return false
  }
  text(value.policyId, `${path}/policyId`, issues)
  if (value.expiresAt !== null && !timestamp(value.expiresAt, `${path}/expiresAt`, issues)) return true
  if (typeof value.expiresAt === 'number' && createdAt !== undefined && value.expiresAt < createdAt) {
    issue(issues, `${path}/expiresAt`, 'INVALID_ORDER', 'expiry must not precede version creation')
  }
  if (typeof value.legalHold !== 'boolean') issue(issues, `${path}/legalHold`, 'REQUIRED', 'legalHold must be boolean')
  return true
}

export function validateHistoryAuthor(input: unknown): HistoryResult<HistoryAuthor> {
  const issues: HistoryIssue[] = []
  validateAuthorInto(input, '/author', issues)
  return issues.length ? { ok: false, issues } : { ok: true, value: structuredClone(input as HistoryAuthor) }
}

export function validateHistoryVersionInfo(input: unknown): HistoryResult<HistoryVersionInfo> {
  const issues: HistoryIssue[] = []
  if (!record(input)) return { ok: false, issues: [{ path: '/', code: 'REQUIRED', message: 'version must be an object' }] }
  text(input.id, '/id', issues)
  text(input.artifactId, '/artifactId', issues)
  if (!Number.isSafeInteger(input.sequence) || (input.sequence as number) < 1) issue(issues, '/sequence', 'INVALID_NUMBER', 'sequence must be a positive integer')
  const hasCreatedAt = timestamp(input.createdAt, '/createdAt', issues)
  if (!Number.isSafeInteger(input.size) || (input.size as number) < 0) issue(issues, '/size', 'INVALID_NUMBER', 'size must be a non-negative integer')
  text(input.contentType, '/contentType', issues, 256)
  validateAuthorInto(input.author, '/author', issues)
  if (!REASONS.has(input.reason as HistoryChangeReason)) issue(issues, '/reason', 'INVALID_ENUM', 'unsupported history reason')
  if (input.sourceVersionId !== undefined) text(input.sourceVersionId, '/sourceVersionId', issues)
  if (input.reason === 'restore' && input.sourceVersionId === undefined) issue(issues, '/sourceVersionId', 'REQUIRED', 'restore versions require a source version id')
  if (input.reason !== 'restore' && input.sourceVersionId !== undefined) issue(issues, '/sourceVersionId', 'INVALID_VALUE', 'only restore versions may have a source version id')
  if (input.description !== undefined && (typeof input.description !== 'string' || input.description.length > MAX_TEXT)) issue(issues, '/description', 'INVALID_VALUE', `description must not exceed ${MAX_TEXT} characters`)
  validateRetentionInto(input.retention, '/retention', issues, hasCreatedAt ? input.createdAt as number : undefined)
  return issues.length ? { ok: false, issues } : { ok: true, value: structuredClone(input as unknown as HistoryVersionInfo) }
}

function uniqueTextArray(value: unknown, path: string, issues: HistoryIssue[]): void {
  if (!Array.isArray(value) || value.length === 0 || value.length > 100 || value.some((entry) => typeof entry !== 'string' || entry.length === 0 || entry.length > MAX_ID)) {
    issue(issues, path, 'INVALID_VALUE', 'must be a non-empty array of 1-512 character strings')
    return
  }
  if (new Set(value).size !== value.length) issue(issues, path, 'DUPLICATE', 'values must be unique')
}

function validateFilter(value: unknown, path: string, issues: HistoryIssue[]): void {
  if (!record(value)) {
    issue(issues, path, 'INVALID_VALUE', 'filter must be an object')
    return
  }
  if (value.authorIds !== undefined) uniqueTextArray(value.authorIds, `${path}/authorIds`, issues)
  if (value.authorKinds !== undefined) {
    uniqueTextArray(value.authorKinds, `${path}/authorKinds`, issues)
    if (Array.isArray(value.authorKinds) && value.authorKinds.some((kind) => !AUTHOR_KINDS.has(kind as HistoryAuthorKind))) issue(issues, `${path}/authorKinds`, 'INVALID_ENUM', 'unsupported author kind')
  }
  if (value.reasons !== undefined) {
    uniqueTextArray(value.reasons, `${path}/reasons`, issues)
    if (Array.isArray(value.reasons) && value.reasons.some((reason) => !REASONS.has(reason as HistoryChangeReason))) issue(issues, `${path}/reasons`, 'INVALID_ENUM', 'unsupported history reason')
  }
  const afterOk = value.createdAfter === undefined || timestamp(value.createdAfter, `${path}/createdAfter`, issues)
  const beforeOk = value.createdBefore === undefined || timestamp(value.createdBefore, `${path}/createdBefore`, issues)
  if (afterOk && beforeOk && typeof value.createdAfter === 'number' && typeof value.createdBefore === 'number' && value.createdAfter > value.createdBefore) {
    issue(issues, path, 'INVALID_ORDER', 'createdAfter must not exceed createdBefore')
  }
  if (value.sourceVersionId !== undefined) text(value.sourceVersionId, `${path}/sourceVersionId`, issues)
}

export function validateHistoryListOptions(input: unknown): HistoryResult<HistoryListOptions> {
  const issues: HistoryIssue[] = []
  if (!record(input)) return { ok: false, issues: [{ path: '/', code: 'INVALID_VALUE', message: 'list options must be an object' }] }
  if (input.cursor !== undefined) text(input.cursor, '/cursor', issues, 4096)
  if (input.limit !== undefined && (!Number.isSafeInteger(input.limit) || (input.limit as number) < 1 || (input.limit as number) > 200)) issue(issues, '/limit', 'INVALID_NUMBER', 'limit must be an integer within 1..200')
  if (input.filter !== undefined) validateFilter(input.filter, '/filter', issues)
  if (issues.length) return { ok: false, issues }
  const { signal, ...data } = input as unknown as HistoryListOptions
  return { ok: true, value: { ...structuredClone(data), ...(signal === undefined ? {} : { signal }) } }
}

export function validateHistoryListPage(input: unknown, artifactId: string): HistoryResult<HistoryListPage> {
  const issues: HistoryIssue[] = []
  if (!record(input) || !Array.isArray(input.versions)) return { ok: false, issues: [{ path: '/versions', code: 'REQUIRED', message: 'host response must contain a versions array' }] }
  const versions: HistoryVersionInfo[] = []
  const ids = new Set<string>()
  for (const [index, value] of input.versions.entries()) {
    const result = validateHistoryVersionInfo(value)
    if (!result.ok) {
      issues.push(...result.issues.map((entry) => ({ ...entry, path: `/versions/${index}${entry.path}` })))
      continue
    }
    if (result.value.artifactId !== artifactId) issue(issues, `/versions/${index}/artifactId`, 'INVALID_VALUE', 'host returned a version for another artifact')
    if (ids.has(result.value.id)) issue(issues, `/versions/${index}/id`, 'DUPLICATE', 'host returned a duplicate version')
    ids.add(result.value.id)
    const previous = versions.at(-1)
    if (previous && (result.value.createdAt > previous.createdAt || result.value.sequence >= previous.sequence)) issue(issues, `/versions/${index}`, 'INVALID_ORDER', 'versions must be newest first')
    versions.push(result.value)
  }
  if (input.nextCursor !== undefined && (typeof input.nextCursor !== 'string' || input.nextCursor.length === 0 || input.nextCursor.length > 4096)) issue(issues, '/nextCursor', 'INVALID_VALUE', 'next cursor must contain 1-4096 characters')
  return issues.length ? { ok: false, issues } : { ok: true, value: { versions, ...(input.nextCursor === undefined ? {} : { nextCursor: input.nextCursor as string }) } }
}

export function historyVersionMatches(version: HistoryVersionInfo, filter: HistoryVersionFilter | undefined): boolean {
  if (!filter) return true
  return (!filter.authorIds || filter.authorIds.includes(version.author.id))
    && (!filter.authorKinds || filter.authorKinds.includes(version.author.kind))
    && (!filter.reasons || filter.reasons.includes(version.reason))
    && (filter.createdAfter === undefined || version.createdAt >= filter.createdAfter)
    && (filter.createdBefore === undefined || version.createdAt <= filter.createdBefore)
    && (filter.sourceVersionId === undefined || version.sourceVersionId === filter.sourceVersionId)
}

export function validateRequestedRetention(input: unknown): HistoryResult<Partial<HistoryRetention>> {
  const issues: HistoryIssue[] = []
  if (!record(input)) return { ok: false, issues: [{ path: '/retention', code: 'INVALID_VALUE', message: 'retention request must be an object' }] }
  if (input.policyId !== undefined) text(input.policyId, '/retention/policyId', issues)
  if (input.expiresAt !== undefined && input.expiresAt !== null) timestamp(input.expiresAt, '/retention/expiresAt', issues)
  if (input.legalHold !== undefined && typeof input.legalHold !== 'boolean') issue(issues, '/retention/legalHold', 'INVALID_VALUE', 'legalHold must be boolean')
  return issues.length ? { ok: false, issues } : { ok: true, value: structuredClone(input as Partial<HistoryRetention>) }
}
