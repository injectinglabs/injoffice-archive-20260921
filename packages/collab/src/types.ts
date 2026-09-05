import type { LogState, OpEntry, OpRecord } from './sync'

/** Wire alias: the room's op-log position as join returns it. */
export type LogStateWire = LogState

// @injoffice/collab — wire + model types. The host's collaboration server
// (the "reference server") owns rooms and op ordering; this package owns
// what a "selection" means and how peers render. Everything here is plain
// JSON, and the transport is host-supplied — any server that speaks this
// small contract works.

/** One peer in a room — the local user ("self") or a remote one. */
export interface PeerInfo<TSelection = SheetSelection> {
  client_id: string
  user_id: string
  name: string
  color: string
  /** Opaque to the gateway; parsed by the editor that owns the room. */
  selection?: TSelection | null
  joined_at: number
}

/**
 * A peer's selection on a sheet. Ranges are [startRow, startColumn, endRow,
 * endColumn], 0-based inclusive (Univer's IRange order). `active` is the
 * anchor cell when known.
 */
export interface SheetSelection {
  sheet: string
  ranges: number[][]
  active?: [number, number]
  /** Whether the active cell is selected or has its editor open. */
  mode?: 'selecting' | 'editing'
  /** Ephemeral, uncommitted plain-text value while `mode` is `editing`. */
  draft?: string
}

/** A collaborator's current focus in a DeckSpec-backed presentation. */
export interface DeckSelection {
  /** Stable SlideSpec id, never a positional index. */
  slideId: string
  /** Stable compiled shape key when the collaborator is editing one. */
  shapeKey?: string | null
}

/** What followers learn when the file on disk changed (gateway FileChange). */
export interface FileChange {
  room: string
  path: string
  author: 'user' | 'agent'
  action: 'saved' | 'print-setup' | 'delivered' | string
  origin?: string
  version?: string
  mtime: number
  size: number
  user_id?: string
  user_name?: string
  /** Op-log seq the bytes on disk reflect (Stage B); 0 from older gateways. */
  saved_seq?: number
  /** The write replaced the file wholesale (agent delivery): log cleared, reload. */
  reset?: boolean
}

/** collab.join response. log is absent on a Stage-A-only gateway. */
export interface JoinResult<TSelection = SheetSelection> {
  room: string
  self: PeerInfo<TSelection>
  peers: PeerInfo<TSelection>[]
  file: { version?: string; mtime: number; size: number }
  log?: LogState
}

/** Events the gateway pushes to room peers. */
export type CollabEvent =
  | { event: 'collab.peer.joined'; payload: { room: string; peer: PeerInfo } }
  | { event: 'collab.peer.left'; payload: { room: string; client_id: string } }
  | { event: 'collab.presence'; payload: { room: string; client_id: string; selection: unknown } }
  | { event: 'collab.file.changed'; payload: FileChange }
  | { event: 'collab.op'; payload: OpEntry }

export const COLLAB_EVENTS = new Set<string>([
  'collab.peer.joined',
  'collab.peer.left',
  'collab.presence',
  'collab.file.changed',
  'collab.op',
])

/**
 * The host supplies the socket. Methods map 1:1 onto the gateway RPCs;
 * onEvent must deliver only collab.* frames (filter on COLLAB_EVENTS);
 * onReconnect fires after the socket re-established itself, at which point
 * the manager re-joins (the gateway gave the new socket a new client id).
 */
export interface CollabTransport<TSelection = SheetSelection> {
  join(path: string, name: string): Promise<JoinResult<TSelection>>
  leave(path: string): Promise<void>
  presence(path: string, selection: TSelection): Promise<void>
  onEvent(handler: (ev: CollabEvent) => void): () => void
  onReconnect(handler: () => void): () => void
  /**
   * Ship ordered edits; resolves with the assigned seq. baseSeq is the head
   * these ops were generated against — the gateway rejects a stale base
   * (throw an Error whose message contains 'STALE_BASE') so the log stays
   * linear; the manager catches up and the queue resubmits.
   */
  opSubmit?(path: string, ops: OpRecord[], baseSeq: number): Promise<number>
  /** Stage B: replay entries after a seq (reset=true: reload the file first). */
  opSince?(path: string, sinceSeq: number): Promise<{ ops: OpEntry[]; head: number; reset: boolean }>
}
