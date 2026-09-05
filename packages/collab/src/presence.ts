import { parseSelection } from './selection'
import type { OpEntry } from './sync'
import type { CollabEvent, FileChange, JoinResult, PeerInfo, SheetSelection } from './types'

// PresenceState — the pure, transport-free model of one room as seen by one
// client. The manager feeds it join results and events; React reads peers().
// Kept free of Univer/DOM so it's unit-testable and reusable by the docs
// editor later.

export class PresenceState<TSelection = SheetSelection> {
  private room = ''
  private selfId = ''
  private self: PeerInfo<TSelection> | null = null
  private readonly remote = new Map<string, PeerInfo<TSelection>>()
  private readonly parse: (raw: unknown) => TSelection | null

  constructor(parse: (raw: unknown) => TSelection | null = parseSelection as unknown as (raw: unknown) => TSelection | null) {
    this.parse = parse
  }

  get roomKey(): string {
    return this.room
  }

  get me(): PeerInfo<TSelection> | null {
    return this.self
  }

  get clientId(): string {
    return this.selfId
  }

  /** Reset from a (re)join. Remote peers are replaced wholesale. */
  applyJoin(res: JoinResult<TSelection>): void {
    if (this.selfId && this.selfId !== res.self.client_id) this.formerIds.add(this.selfId)
    this.room = res.room
    this.self = res.self
    this.selfId = res.self.client_id
    this.remote.clear()
    for (const p of res.peers) {
      if (p.client_id === this.selfId) continue
      this.remote.set(p.client_id, { ...p, selection: this.parse(p.selection) })
    }
  }

  clear(): void {
    this.room = ''
    this.self = null
    this.selfId = ''
    this.remote.clear()
  }

  /**
   * Fold one gateway event in. Returns what changed so the manager can do
   * the minimum: 'peers' (list changed), 'selection' (one peer moved),
   * 'file' (disk changed; the change is returned), or null (not ours).
   */
  apply(
    ev: CollabEvent,
  ): { kind: 'peers' } | { kind: 'selection'; peer: PeerInfo<TSelection> } | { kind: 'file'; change: FileChange } | { kind: 'op'; entry: OpEntry } | null {
    if (!this.room || ev.payload.room !== this.room) return null
    switch (ev.event) {
      case 'collab.peer.joined': {
        const p = ev.payload.peer
        if (p.client_id === this.selfId) return null
        this.remote.set(p.client_id, { ...p, selection: this.parse(p.selection) })
        return { kind: 'peers' }
      }
      case 'collab.peer.left': {
        if (!this.remote.delete(ev.payload.client_id)) return null
        return { kind: 'peers' }
      }
      case 'collab.presence': {
        const p = this.remote.get(ev.payload.client_id)
        if (!p) return null
        p.selection = this.parse(ev.payload.selection)
        return { kind: 'selection', peer: p }
      }
      case 'collab.file.changed':
        if (ev.payload.origin && ev.payload.origin === this.selfId) return null
        return { kind: 'file', change: ev.payload }
      case 'collab.op':
        if (ev.payload.client_id === this.selfId) return null
        return { kind: 'op', entry: ev.payload }
      default:
        return null
    }
  }

  /** Rewrite every stored remote selection (structural shift). */
  shiftSelections(fn: (sel: NonNullable<PeerInfo<TSelection>['selection']>) => NonNullable<PeerInfo<TSelection>['selection']>): boolean {
    let changed = false
    for (const p of this.remote.values()) {
      if (!p.selection) continue
      const next = fn(p.selection)
      if (next !== p.selection) {
        p.selection = next
        changed = true
      }
    }
    return changed
  }

  /** Remote peers, oldest first. */
  peers(): PeerInfo<TSelection>[] {
    return [...this.remote.values()].sort((a, b) => a.joined_at - b.joined_at)
  }

  /** Has this client id ever been us (reconnects hand out new ids)? */
  private readonly formerIds = new Set<string>()

  isSelf(clientId: string): boolean {
    return clientId === this.selfId || this.formerIds.has(clientId)
  }

  /** Display name for a client id (used by the file-changed banner). */
  nameOf(clientId: string | undefined): string | null {
    if (!clientId) return null
    return this.remote.get(clientId)?.name ?? null
  }
}
