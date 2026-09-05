import { DeckSyncEngine, type DeckSyncHooks, type SyncDeck } from './deckSync'
import { PresenceState } from './presence'
import type { CollabTransport, DeckSelection, FileChange, LogStateWire, PeerInfo } from './types'

function parseDeckSelection(raw: unknown): DeckSelection | null {
  if (!raw || typeof raw !== 'object') return null
  const value = raw as Record<string, unknown>
  if (typeof value.slideId !== 'string') return null
  return { slideId: value.slideId, shapeKey: typeof value.shapeKey === 'string' ? value.shapeKey : null }
}

/** Presence and ordered co-editing for a DeckSpec JSON file. */
export class DeckPresenceManager<TDeck extends SyncDeck> {
  readonly state = new PresenceState<DeckSelection>(parseDeckSelection)
  private readonly changes = new Set<() => void>()
  private readonly fileChanges = new Set<(change: FileChange) => void>()
  private readonly disposables: Array<() => void> = []
  private started = false
  private joined = false
  private logState: LogStateWire | null = null
  private engine: DeckSyncEngine<TDeck> | null = null

  constructor(private readonly transport: CollabTransport<DeckSelection>, private readonly opts: { path: string; name: string }) {}

  get clientId(): string | null { return this.joined ? this.state.clientId || null : null }
  get log(): LogStateWire | null { return this.logState }
  get appliedSeq(): number { return this.engine?.applied ?? 0 }
  get syncing(): boolean { return this.engine?.syncing ?? false }
  peers(): PeerInfo<DeckSelection>[] { return this.state.peers() }
  onChange(cb: () => void): () => void { this.changes.add(cb); return () => this.changes.delete(cb) }
  onFileChanged(cb: (change: FileChange) => void): () => void { this.fileChanges.add(cb); return () => this.fileChanges.delete(cb) }

  publish(selection: DeckSelection): void {
    if (this.joined) void this.transport.presence(this.opts.path, selection).catch(() => undefined)
  }

  async start(): Promise<void> {
    if (this.started) return
    this.started = true
    this.disposables.push(this.transport.onEvent((event) => {
      const result = this.state.apply(event)
      if (!result) return
      if (result.kind === 'peers' || result.kind === 'selection') this.notify()
      else if (result.kind === 'op') this.engine?.receiveEntry(result.entry.seq, result.entry.ops as unknown[])
      else if (result.kind === 'file') {
        if (this.engine && !result.change.reset && (result.change.saved_seq ?? 0) <= this.engine.applied) return
        for (const listener of this.fileChanges) listener(result.change)
      }
    }))
    this.disposables.push(this.transport.onReconnect(() => { void this.join().then(() => this.engine?.resume()) }))
    await this.join()
  }

  async startSync(initial: TDeck, hooks: DeckSyncHooks<TDeck>): Promise<DeckSyncEngine<TDeck> | null> {
    if (!this.joined || !this.logState || !this.transport.opSubmit || !this.transport.opSince) return null
    const engine = new DeckSyncEngine(this.transport, this.opts.path, initial, hooks)
    this.engine = engine
    await engine.bootstrap(this.logState.saved_seq)
    return engine
  }

  stop(): void {
    if (!this.started) return
    this.started = false
    this.engine?.stop(); this.engine = null
    for (const dispose of this.disposables.splice(0)) dispose()
    if (this.joined) { this.joined = false; void this.transport.leave(this.opts.path).catch(() => undefined) }
    this.state.clear(); this.changes.clear(); this.fileChanges.clear()
  }

  private async join(): Promise<void> {
    try {
      const result = await this.transport.join(this.opts.path, this.opts.name)
      if (!this.started) return
      this.state.applyJoin(result)
      this.logState = result.log ?? null
      this.joined = true
      this.notify()
    } catch { this.joined = false }
  }

  private notify(): void { for (const listener of this.changes) listener() }
}
