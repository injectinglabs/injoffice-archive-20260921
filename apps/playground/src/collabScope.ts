export const COLLAB_COPY = {
  proof: 'This page mounts a real Univer sheet and sends cell mutations, ordered replay, reconnect catch-up, and selection presence through injoffice-server over HTTP and server-sent events.',
  security: 'The demo has no authentication, authorization, tenant isolation, or rate limiting. Keep the sidecar bound to localhost; do not expose it to a network.',
  completion: 'The scoped native-office v3 checklist is complete. Broader format fidelity, editor integration, and production collaboration coverage remain partial.',
} as const

export const ARTIFACT_PARAM = 'artifact'
export const FORMAT_PARAM = 'format'
export const COLLAB_FORMATS = ['sheets', 'slides', 'docs', 'pdf'] as const
export type CollabFormat = (typeof COLLAB_FORMATS)[number]
export const DEFAULT_COLLAB_FORMAT: CollabFormat = 'sheets'

const FORMAT_SET = new Set<string>(COLLAB_FORMATS)

export function parseCollabFormat(value: string | null | undefined): CollabFormat {
  const next = value?.trim().toLowerCase() ?? ''
  return FORMAT_SET.has(next) ? (next as CollabFormat) : DEFAULT_COLLAB_FORMAT
}

/** Read `artifact` and `format` from a search string, hash query, or href. Search wins when both exist. */
export function parseCollabQuery(input: string): { artifact: string; format: CollabFormat } {
  const url = new URL(input, 'http://collab.invalid')
  const hash = url.hash.startsWith('#') ? url.hash.slice(1) : url.hash
  const hashQuery = hash.includes('?') ? hash.slice(hash.indexOf('?') + 1) : ''
  const merged = new URLSearchParams(hashQuery)
  for (const [key, value] of url.searchParams) merged.set(key, value)
  return {
    artifact: merged.get(ARTIFACT_PARAM)?.trim() ?? '',
    format: parseCollabFormat(merged.get(FORMAT_PARAM)),
  }
}

/** Write `artifact` and `format` onto the search string without dropping the other, and strip them from the hash query. */
export function applyCollabQuery(href: string, next: { artifact: string; format: CollabFormat }): string {
  const url = new URL(href, 'http://collab.invalid')
  if (next.artifact) url.searchParams.set(ARTIFACT_PARAM, next.artifact)
  else url.searchParams.delete(ARTIFACT_PARAM)
  url.searchParams.set(FORMAT_PARAM, next.format)
  if (url.hash.includes('?')) {
    const raw = url.hash.slice(1)
    const qi = raw.indexOf('?')
    const path = raw.slice(0, qi)
    const hashParams = new URLSearchParams(raw.slice(qi + 1))
    hashParams.delete(ARTIFACT_PARAM)
    hashParams.delete(FORMAT_PARAM)
    const rest = hashParams.toString()
    url.hash = rest ? `#${path}?${rest}` : `#${path}`
  }
  return `${url.pathname}${url.search}${url.hash}`
}

export function writeCollabQuery(next: { artifact: string; format: CollabFormat }) {
  window.history.replaceState({}, '', applyCollabQuery(window.location.href, next))
}
