import { DocSyncEngine, type DocSyncHooks } from './docSync'
import { PresenceState } from './presence'
import type { CollabTransport, FileChange, LogStateWire, PeerInfo } from './types'

export interface DocSelection {
  from: number
  to: number
}

function parseDocSelection(raw: unknown): DocSelection | null {
  if (!raw || typeof raw !== 'object') return null
  const value = raw as Record<string, unknown>
  if (typeof value.from !== 'number' || typeof value.to !== 'number') return null
  if (!Number.isInteger(value.from) || !Number.isInteger(value.to) || value.from < 0 || value.to < 0) return null
  return { from: value.from, to: value.to }
}

// Rooms are path-keyed.
// Ops are opaque and the host supplies editor OT.
export class DocPresenceManager {
  readonly state = new PresenceState<DocSelection>(parseDocSelection)
  private readonly transport: CollabTransport<DocSelection>
  private readonly path: string
  private readonly name: string
  private readonly changeListeners = new Set<() => void>()
  private readonly fileListeners = new Set<(c: FileChange) => void>()
  private readonly disposables: Array<() => void> = []
  private started = false
  private joined = false
  private logState: LogStateWire | null = null
  private engine: DocSyncEngine | null = null

  constructor(transport: CollabTransport<DocSelection>, opts: { path: string; name: string }) {
    this.transport = transport
    this.path = opts.path
    this.name = opts.name
  }

  /** Our server client id once joined — the collab_origin of our own PUT. */
  get clientId(): string | null {
    return this.joined ? this.state.clientId || null : null
  }

  /** The room's op-log position at join (null on a presence-only server). */
  get log(): LogStateWire | null {
    return this.logState
  }

  /** Last op-log entry applied by the sync engine (Save's collab_seq). */
  get appliedSeq(): number {
    return this.engine?.applied ?? 0
  }

  /** Whether live co-editing is running. */
  get syncing(): boolean {
    return this.engine !== null
  }

  /**
   * Start live co-editing (D3): the host loads the document at
   * log.saved_seq FIRST, then calls this — the engine replays newer
   * entries and keeps the editor in the room's total order from there.
   * No-op (returns null) when the server has no op log.
   */
  async startSync(hooks: DocSyncHooks): Promise<DocSyncEngine | null> {
    if (!this.joined || !this.logState || !this.transport.opSubmit || !this.transport.opSince) return null
    const engine = new DocSyncEngine(this.transport as unknown as CollabTransport, this.path, hooks, this.state.clientId || null)
    this.engine = engine
    await engine.bootstrap(this.logState.saved_seq)
    // The editor is usable while bootstrap is in flight. Any steps authored
    // in that window were recorded by ProseMirror, but its onUpdate callback
    // could not submit through an engine that had not attached yet. Flush
    // once after catch-up so those first keystrokes are not stranded.
    engine.submitSoon()
    return engine
  }

  /**
   * Re-establish live co-editing after the server RESET the room's op log
   * (an agent delivery replaced the file wholesale): the old engine's seqs
   * are void, so it is stopped, the room re-joined for the fresh log state,
   * and a new engine bootstrapped from the new saved_seq. The host calls
   * this AFTER it has brought its document up to date with the new file
   * (e.g. an in-place merge), so there is nothing to replay.
   */
  async restartSync(hooks: DocSyncHooks): Promise<DocSyncEngine | null> {
    this.engine?.stop()
    this.engine = null
    await this.join()
    return this.startSync(hooks)
  }

  peers(): PeerInfo<DocSelection>[] {
    return this.state.peers()
  }

  publish(selection: DocSelection): void {
    if (this.joined) void this.transport.presence(this.path, selection).catch(() => undefined)
  }

  onChange(cb: () => void): () => void {
    this.changeListeners.add(cb)
    return () => this.changeListeners.delete(cb)
  }

  /** The file on disk changed under this editor — reload or banner. */
  onFileChanged(cb: (c: FileChange) => void): () => void {
    this.fileListeners.add(cb)
    return () => this.fileListeners.delete(cb)
  }

  async start(): Promise<void> {
    if (this.started) return
    this.started = true
    this.disposables.push(
      this.transport.onEvent((ev) => {
        const res = this.state.apply(ev)
        if (!res) return
        if (res.kind === 'peers' || res.kind === 'selection') for (const cb of this.changeListeners) cb()
        else if (res.kind === 'op') this.engine?.receiveEntry(res.entry.seq, res.entry.ops as unknown[], res.entry.client_id)
        else if (res.kind === 'file') {
          // While co-editing, a peer's save we are already at (or ahead of)
          // is not news — our doc contains those ops. Reset means the file
          // was replaced wholesale (agent delivery): always surface.
          const savedSeq = res.change.saved_seq ?? 0
          if (this.engine && !res.change.reset && savedSeq <= this.engine.applied) return
          for (const cb of this.fileListeners) cb(res.change)
        }
      }),
    )
    this.disposables.push(
      this.transport.onReconnect(() => {
        void this.join().then(() => this.engine?.resume())
      }),
    )
    await this.join()
  }

  stop(): void {
    if (!this.started) return
    this.started = false
    this.engine?.stop()
    this.engine = null
    for (const d of this.disposables.splice(0)) d()
    if (this.joined) {
      this.joined = false
      void this.transport.leave(this.path).catch(() => undefined)
    }
    this.state.clear()
    this.changeListeners.clear()
    this.fileListeners.clear()
  }

  private async join(): Promise<void> {
    try {
      const res = await this.transport.join(this.path, this.name)
      if (!this.started) return
      this.state.applyJoin(res)
      this.logState = res.log ?? null
      this.joined = true
      for (const cb of this.changeListeners) cb()
    } catch {
      // Old server or socket down: the editor works exactly as before,
      // just without presence. Reconnect retries.
      this.joined = false
    }
  }
}
