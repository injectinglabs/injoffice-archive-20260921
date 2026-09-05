import type {
  CollabEvent,
  CollabTransport,
  JoinResult,
  PeerInfo,
  SheetSelection,
} from '../../../packages/collab/src/types.js'
import type { OpEntry, OpRecord } from '../../../packages/collab/src/sync.js'

export type BrowserCollabActivity = {
  id: number
  actor: string
  detail: string
  kind: 'join' | 'leave' | 'presence' | 'operation'
  seq?: number
}

export type BrowserCollabSnapshot = {
  peers: number
  head: number
  activity: BrowserCollabActivity[]
}

type ClientProfile = {
  userId: string
  color: string
}

type BrowserClient = ClientProfile & {
  id: string
  closed: boolean
  events: Set<(event: CollabEvent) => void>
  reconnects: Set<() => void>
  rooms: Set<string>
}

type BrowserRoom = {
  path: string
  createdAt: number
  head: number
  entries: OpEntry[]
  peers: Map<string, PeerInfo<unknown>>
}

const clone = <T,>(value: T): T => JSON.parse(JSON.stringify(value)) as T

export function describePresence(selection: unknown): string {
  if (!selection || typeof selection !== 'object') return 'moved the shared selection'
  const value = selection as Record<string, unknown>
  if (value.mode === 'editing') {
    return typeof value.draft === 'string'
      ? `typing “${String(value.draft).slice(0, 36) || '…'}”`
      : 'started editing a cell'
  }
  if (typeof value.from === 'number' && typeof value.to === 'number') {
    return value.from === value.to ? `caret at ${value.from}` : `selected ${value.from}–${value.to}`
  }
  if (typeof value.slideId === 'string') {
    return typeof value.shapeKey === 'string' && value.shapeKey
      ? `editing a shape on ${value.slideId}`
      : `viewing slide ${value.slideId}`
  }
  if (typeof value.page === 'number') return `on page ${value.page}`
  return 'moved the shared selection'
}

/**
 * Small, deterministic browser implementation of the public collaboration
 * transport contract. It is intentionally demo-only: state lives for the
 * lifetime of this object and never touches an Office package or a network.
 */
export class BrowserCollabHub {
  private readonly rooms = new Map<string, BrowserRoom>()
  private readonly clients = new Map<string, BrowserClient>()
  private readonly listeners = new Set<() => void>()
  private nextClient = 1
  private nextActivity = 1
  private snapshot: BrowserCollabSnapshot = { peers: 0, head: 0, activity: [] }

  readonly subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  readonly getSnapshot = (): BrowserCollabSnapshot => this.snapshot

  connect<TSelection = SheetSelection>(profile: ClientProfile): CollabTransport<TSelection> & { close(): void } {
    const client: BrowserClient = {
      ...profile,
      id: `browser-client-${this.nextClient++}`,
      closed: false,
      events: new Set(),
      reconnects: new Set(),
      rooms: new Set(),
    }
    this.clients.set(client.id, client)

    const requireRoom = (path: string): BrowserRoom => {
      const room = this.rooms.get(path)
      if (!room || !room.peers.has(client.id)) throw new Error('NOT_IN_ROOM')
      return room
    }

    const leave = async (path: string) => {
      const room = this.rooms.get(path)
      if (!room || !room.peers.has(client.id)) return
      const peer = room.peers.get(client.id)!
      room.peers.delete(client.id)
      client.rooms.delete(path)
      this.broadcast(room, client.id, {
        event: 'collab.peer.left',
        payload: { room: this.roomKey(path), client_id: client.id },
      })
      this.record('leave', peer.name, 'left the browser room')
    }

    return {
      join: async (path, name): Promise<JoinResult<TSelection>> => {
        if (client.closed) throw new Error('SIMULATOR_CLOSED')
        const room = this.room(path)
        const existing = room.peers.get(client.id)
        const self: PeerInfo<TSelection> = {
          client_id: client.id,
          user_id: client.userId,
          name: name.trim() || client.userId,
          color: client.color,
          selection: (existing?.selection ?? null) as TSelection | null,
          joined_at: existing?.joined_at ?? Date.now(),
        }
        const peers = [...room.peers.values()]
          .filter((peer) => peer.client_id !== client.id)
          .map((peer) => clone(peer) as PeerInfo<TSelection>)
        room.peers.set(client.id, self)
        client.rooms.add(path)
        if (!existing) {
          this.broadcast(room, client.id, {
            event: 'collab.peer.joined',
            payload: { room: this.roomKey(path), peer: clone(self) as PeerInfo },
          })
          this.record('join', self.name, 'joined the browser room')
        }
        return {
          room: this.roomKey(path),
          self: clone(self),
          peers,
          file: { version: 'browser-snapshot-v1', mtime: room.createdAt, size: 0 },
          log: { seq: room.head, saved_seq: 0 },
        }
      },
      leave,
      presence: async (path, selection: TSelection) => {
        const room = requireRoom(path)
        const peer = room.peers.get(client.id)!
        peer.selection = clone(selection)
        this.broadcast(room, client.id, {
          event: 'collab.presence',
          payload: { room: this.roomKey(path), client_id: client.id, selection: clone(selection) },
        })
        this.record('presence', peer.name, describePresence(selection))
      },
      onEvent: (handler) => {
        client.events.add(handler)
        return () => client.events.delete(handler)
      },
      onReconnect: (handler) => {
        client.reconnects.add(handler)
        return () => client.reconnects.delete(handler)
      },
      opSubmit: async (path, ops: OpRecord[], baseSeq: number) => {
        const room = requireRoom(path)
        if (baseSeq !== room.head) throw new Error('STALE_BASE')
        const peer = room.peers.get(client.id)!
        const entry: OpEntry = {
          room: this.roomKey(path),
          seq: ++room.head,
          client_id: client.id,
          ops: clone(ops),
        }
        room.entries.push(entry)
        this.broadcast(room, client.id, { event: 'collab.op', payload: clone(entry) })
        this.record('operation', peer.name, `${ops.length} mutation${ops.length === 1 ? '' : 's'} committed`, entry.seq)
        return entry.seq
      },
      opSince: async (path, sinceSeq) => {
        const room = requireRoom(path)
        if (sinceSeq < 0 || sinceSeq > room.head) return { ops: [], head: room.head, reset: true }
        return {
          ops: room.entries.filter((entry) => entry.seq > sinceSeq).map(clone),
          head: room.head,
          reset: false,
        }
      },
      close: () => {
        if (client.closed) return
        client.closed = true
        for (const path of [...client.rooms]) void leave(path)
        client.events.clear()
        client.reconnects.clear()
        this.clients.delete(client.id)
      },
    }
  }

  private room(path: string): BrowserRoom {
    let room = this.rooms.get(path)
    if (!room) {
      room = { path, createdAt: Date.now(), head: 0, entries: [], peers: new Map() }
      this.rooms.set(path, room)
    }
    return room
  }

  private roomKey(path: string): string {
    return `browser:${path}`
  }

  private broadcast(room: BrowserRoom, sender: string, event: CollabEvent) {
    for (const peer of room.peers.values()) {
      if (peer.client_id === sender) continue
      const target = this.clients.get(peer.client_id)
      if (!target || target.closed) continue
      for (const handler of target.events) handler(clone(event))
    }
  }

  private record(kind: BrowserCollabActivity['kind'], actor: string, detail: string, seq?: number) {
    const activity = [
      { id: this.nextActivity++, kind, actor, detail, ...(seq === undefined ? {} : { seq }) },
      ...this.snapshot.activity,
    ].slice(0, 8)
    let peers = 0
    let head = 0
    for (const room of this.rooms.values()) {
      peers += room.peers.size
      head = Math.max(head, room.head)
    }
    this.snapshot = { peers, head, activity }
    for (const listener of this.listeners) listener()
  }
}

export function createBrowserCollabHub(): BrowserCollabHub {
  return new BrowserCollabHub()
}
