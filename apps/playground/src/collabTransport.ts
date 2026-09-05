import { COLLAB_EVENTS, type CollabEvent, type CollabTransport, type JoinResult } from '../../../packages/collab/src/types.js'
import type { OpEntry, OpRecord } from '../../../packages/collab/src/sync.js'

export type HttpCollabTransport = CollabTransport & {
  sessionId: string
  clientId: string
  close(): void
}

type SessionResponse = { session_id: string; user_id: string; client_id: string }

const eventsTimeoutMs = 5000

async function readError(res: Response): Promise<string> {
  const text = await res.text()
  try {
    const parsed = JSON.parse(text) as { error?: string }
    if (parsed.error) return parsed.error
  } catch {
    /* keep raw body */
  }
  return text || res.statusText
}

async function mintSession(root: string): Promise<SessionResponse> {
  let opened: Response
  try {
    opened = await fetch(`${root}/v1/collab/session`, { method: 'POST' })
  } catch {
    throw new Error('collab session failed: start injoffice-server with cd go/injoffice-server && go run ./cmd/injoffice-server')
  }
  if (!opened.ok) {
    const detail = await readError(opened)
    throw new Error(`collab session failed (${opened.status}): ${detail}`)
  }
  return (await opened.json()) as SessionResponse
}

function waitOpen(source: EventSource, ms: number): Promise<void> {
  return new Promise((resolve, reject) => {
    if (source.readyState === EventSource.OPEN) {
      resolve()
      return
    }
    const timer = window.setTimeout(() => {
      cleanup()
      reject(new Error('collab events timeout: start injoffice-server with cd go/injoffice-server && go run ./cmd/injoffice-server'))
    }, ms)
    const onOpen = () => {
      cleanup()
      resolve()
    }
    const cleanup = () => {
      window.clearTimeout(timer)
      source.removeEventListener('open', onOpen)
    }
    source.addEventListener('open', onOpen)
  })
}

export async function createHttpCollabTransport(base = ''): Promise<HttpCollabTransport> {
  const root = base.replace(/\/$/, '')
  const eventHandlers = new Set<(ev: CollabEvent) => void>()
  const reconnectHandlers = new Set<() => void>()
  let closed = false
  let reminting = false
  let session = await mintSession(root)
  let source: EventSource | null = null

  const rpc = {
    session_id: session.session_id,
  }

  const dispatch = (message: MessageEvent) => {
    let frame: CollabEvent
    try {
      frame = JSON.parse(message.data) as CollabEvent
    } catch {
      return
    }
    if (!COLLAB_EVENTS.has(frame.event)) return
    for (const handler of eventHandlers) handler(frame)
  }

  const fireReconnect = () => {
    if (closed) return
    for (const handler of reconnectHandlers) handler()
  }

  const bindErrorHandler = () => {
    if (!source) return
    source.onerror = () => {
      if (closed || reminting) return
      void probeAndReconnect()
    }
  }

  const attachSource = async (next: SessionResponse) => {
    source?.close()
    session = next
    rpc.session_id = next.session_id
    const events = new EventSource(`${root}/v1/collab/events?session_id=${encodeURIComponent(next.session_id)}`)
    source = events
    events.onmessage = dispatch
    let firstOpen = true
    events.addEventListener('open', () => {
      if (firstOpen) {
        firstOpen = false
        return
      }
      fireReconnect()
    })
    await waitOpen(events, eventsTimeoutMs)
  }

  const probeAndReconnect = async () => {
    if (closed || reminting) return
    reminting = true
    try {
      const res = await fetch(`${root}/v1/collab/session?session_id=${encodeURIComponent(session.session_id)}`)
      if (res.ok) return
      const next = await mintSession(root)
      await attachSource(next)
      live.sessionId = next.session_id
      live.clientId = next.client_id
      bindErrorHandler()
      fireReconnect()
    } catch {
      /* EventSource retries; a later open still catch-up */
    } finally {
      reminting = false
    }
  }

  await attachSource(session)

  let onPageHide = () => {}

  const postJSON = async <T>(url: string, body: unknown): Promise<T> => {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    })
    const text = await res.text()
    if (!res.ok) {
      let detail = text
      try {
        const parsed = JSON.parse(text) as { error?: string }
        if (parsed.error) detail = parsed.error
      } catch {
        /* keep raw body */
      }
      if (res.status === 401) void probeAndReconnect()
      if (res.status === 409 || /STALE_BASE/i.test(detail)) {
        throw new Error(detail.includes('STALE_BASE') ? detail : 'STALE_BASE')
      }
      throw new Error(detail || `${url} failed (${res.status})`)
    }
    if (!text) return undefined as T
    return JSON.parse(text) as T
  }

  const live: HttpCollabTransport = {
    sessionId: session.session_id,
    clientId: session.client_id,
    async join(path, name) {
      return postJSON<JoinResult>(`${root}/v1/collab/join`, { ...rpc, path, name })
    },
    async leave(path) {
      await postJSON(`${root}/v1/collab/leave`, { ...rpc, path })
    },
    async presence(path, selection) {
      await postJSON(`${root}/v1/collab/presence`, { ...rpc, path, selection })
    },
    onEvent(handler) {
      eventHandlers.add(handler)
      return () => {
        eventHandlers.delete(handler)
      }
    },
    onReconnect(handler) {
      reconnectHandlers.add(handler)
      return () => {
        reconnectHandlers.delete(handler)
      }
    },
    async opSubmit(path, ops: OpRecord[], baseSeq) {
      const got = await postJSON<{ seq: number }>(`${root}/v1/collab/op/submit`, {
        ...rpc,
        path,
        base_seq: baseSeq,
        ops,
      })
      return got.seq
    },
    async opSince(path, sinceSeq) {
      return postJSON<{ ops: OpEntry[]; head: number; reset: boolean }>(`${root}/v1/collab/op/since`, {
        ...rpc,
        path,
        since_seq: sinceSeq,
      })
    },
    close() {
      if (closed) return
      closed = true
      window.removeEventListener('pagehide', onPageHide)
      window.removeEventListener('beforeunload', onPageHide)
      const id = session.session_id
      source?.close()
      source = null
      const closeURL = `${root}/v1/collab/session/close?session_id=${encodeURIComponent(id)}`
      try {
        navigator.sendBeacon(closeURL)
      } catch {
        /* keepalive DELETE below */
      }
      void fetch(`${root}/v1/collab/session?session_id=${encodeURIComponent(id)}`, { method: 'DELETE', keepalive: true })
    },
  }
  onPageHide = () => {
    live.close()
  }
  window.addEventListener('pagehide', onPageHide)
  window.addEventListener('beforeunload', onPageHide)
  bindErrorHandler()
  return live
}
