import { SparklineManager, type CreateSparklineInput } from './manager'
import type { SparklineGroup, SparklineSnapshotV1, SparklineSpec } from './types'

export interface SparklineUndoRecord {
  label: string
  before: SparklineSnapshotV1
  after: SparklineSnapshotV1
}

export interface SparklineUndoSink {
  push(record: SparklineUndoRecord): void
}

export class SparklineHandle {
  constructor(private readonly owner: SparklineCommandController, readonly id: string) {}

  get value(): SparklineSpec | undefined { return this.owner.manager.get(this.id) }
  update(patch: Partial<Omit<SparklineSpec, 'id' | 'groupId'>>): SparklineSpec { return this.owner.update(this.id, patch) }
  remove(): boolean { return this.owner.remove(this.id) }
}

/** Host-neutral lifecycle facade that records complete, atomic inverses. */
export class SparklineCommandController {
  constructor(readonly manager: SparklineManager, private readonly undo?: SparklineUndoSink) {}

  create(input: CreateSparklineInput): SparklineHandle {
    const spec = this.record('Create sparkline', () => this.manager.create(input))
    return new SparklineHandle(this, spec.id)
  }

  getById(id: string): SparklineHandle | null {
    return this.manager.get(id) ? new SparklineHandle(this, id) : null
  }

  list(): SparklineHandle[] {
    return this.manager.list().map((spec) => new SparklineHandle(this, spec.id))
  }

  update(id: string, patch: Partial<Omit<SparklineSpec, 'id' | 'groupId'>>): SparklineSpec {
    return this.record('Update sparkline', () => this.manager.update(id, patch))
  }

  remove(id: string): boolean {
    return this.record('Remove sparkline', () => this.manager.remove(id))
  }

  group(memberIds: readonly string[], id?: string): SparklineGroup {
    return this.record('Group sparklines', () => this.manager.group(memberIds, id))
  }

  ungroup(memberIds: readonly string[]): void {
    this.record('Ungroup sparklines', () => this.manager.ungroup(memberIds))
  }

  restore(snapshot: SparklineSnapshotV1): boolean {
    try {
      this.manager.hydrate(snapshot)
      return true
    } catch {
      return false
    }
  }

  private record<T>(label: string, mutation: () => T): T {
    const before = this.manager.serialize()
    const value = mutation()
    const after = this.manager.serialize()
    if (JSON.stringify(before) !== JSON.stringify(after)) this.undo?.push({ label, before, after })
    return value
  }
}
