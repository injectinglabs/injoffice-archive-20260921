import { AgentToolsError, type AgentArtifactIdentity, type JsonValue } from '@injoffice/agent-tools'

const encoder = new TextEncoder()

function canonical(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value)
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`
  const object = value as Record<string, unknown>
  return `{${Object.keys(object).sort().map((key) => `${JSON.stringify(key)}:${canonical(object[key])}`).join(',')}}`
}

export async function fingerprintBytes(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes as Uint8Array<ArrayBuffer>)
  return `sha256:${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`
}

export async function fingerprintJSON(value: unknown): Promise<string> {
  return fingerprintBytes(encoder.encode(canonical(value)))
}

export function errorIssue(error: unknown, path?: string) {
  return {
    code: 'INVALID_OPERATION',
    ...(path ? { path } : {}),
    message: error instanceof Error ? error.message : String(error),
    severity: 'error' as const,
  }
}

export function assertFresh(actual: AgentArtifactIdentity, expected: AgentArtifactIdentity): void {
  if (actual.artifactId !== expected.artifactId) throw new Error('artifact identity changed after the changeset was planned')
  if (actual.revision !== expected.revision || actual.fingerprint !== expected.fingerprint) {
    throw new AgentToolsError('STALE_REVISION', `stale revision: expected ${expected.revision}, received ${actual.revision}`, true)
  }
}

export function toJson(value: unknown): JsonValue {
  return JSON.parse(JSON.stringify(value)) as JsonValue
}

export function cursorOffset(cursor: string | undefined): number {
  if (cursor === undefined) return 0
  if (!/^(0|[1-9][0-9]*)$/.test(cursor)) throw new AgentToolsError('INVALID_ARGUMENT', 'cursor must be a non-negative decimal offset')
  const offset = Number(cursor)
  if (!Number.isSafeInteger(offset)) throw new AgentToolsError('INVALID_ARGUMENT', 'cursor is outside the safe integer range')
  return offset
}

export function boundedObject(
  base: Record<string, JsonValue>,
  key: string,
  items: readonly unknown[],
  offset: number,
  maxItems: number,
  maxBytes: number,
) {
  const selected: JsonValue[] = []
  const empty = { ...base, [key]: selected }
  if (encoder.encode(JSON.stringify(empty)).byteLength > maxBytes) throw new AgentToolsError('LIMIT_EXCEEDED', 'maxBytes is too small for the requested view metadata')
  for (let index = offset; index < items.length && selected.length < maxItems; index += 1) {
    const candidate = [...selected, toJson(items[index])]
    const data = { ...base, [key]: candidate }
    if (encoder.encode(JSON.stringify(data)).byteLength > maxBytes) {
      if (selected.length === 0) throw new AgentToolsError('LIMIT_EXCEEDED', 'the next item cannot fit within maxBytes; request a larger byte bound or a narrower target')
      break
    }
    selected.push(candidate[candidate.length - 1]!)
  }
  const next = offset + selected.length
  const truncated = next < items.length
  return {
    data: { ...base, [key]: selected },
    itemCount: selected.length,
    truncated,
    ...(truncated ? { nextCursor: String(next) } : {}),
  }
}
