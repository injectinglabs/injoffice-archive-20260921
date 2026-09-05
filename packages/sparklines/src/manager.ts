import { extractSparklineValues } from './extract'
import { compileSparklineGeometry, numericExtent } from './geometry'
import type {
  SparklineGeometry, SparklineGroup, SparklineSnapshotV1, SparklineSpec,
  SparklineValueReader, SparklineViewport,
} from './types'
import { validateSparkline, validateSparklineSnapshot } from './validation'

export interface SparklineManagerOptions {
  idFactory?: (kind: 'sparkline' | 'group') => string
}

export interface CreateSparklineInput extends Omit<SparklineSpec, 'id' | 'groupId'> { id?: string }

let sequence = 0
const defaultIdFactory = (kind: 'sparkline' | 'group'): string => `${kind}-${Date.now().toString(36)}-${(++sequence).toString(36)}`
const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T

export class SparklineManager {
  private readonly sparklines = new Map<string, SparklineSpec>()
  private readonly groups = new Map<string, SparklineGroup>()
  private readonly listeners = new Set<(snapshot: SparklineSnapshotV1) => void>()
  private readonly idFactory: (kind: 'sparkline' | 'group') => string

  constructor(options: SparklineManagerOptions = {}) {
    this.idFactory = options.idFactory ?? defaultIdFactory
  }

  onChange(listener: (snapshot: SparklineSnapshotV1) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  create(input: CreateSparklineInput): SparklineSpec {
    const spec: SparklineSpec = { ...clone(input), id: input.id ?? this.idFactory('sparkline') }
    this.add(spec)
    return clone(spec)
  }

  add(spec: SparklineSpec): void {
    const issues = validateSparkline(spec)
    if (issues.length) throw new TypeError(`Invalid sparkline: ${issues.map((issue) => `${issue.path} ${issue.message}`).join('; ')}`)
    if (this.sparklines.has(spec.id)) throw new Error(`Sparkline ${spec.id} already exists`)
    this.assertTargetAvailable(spec.target)
    if (spec.groupId) throw new Error('Use group() to assign a new sparkline to a group')
    this.sparklines.set(spec.id, clone(spec))
    this.emit()
  }

  update(id: string, patch: Partial<Omit<SparklineSpec, 'id' | 'groupId'>>): SparklineSpec {
    const current = this.require(id)
    const next: SparklineSpec = { ...current, ...clone(patch), id, groupId: current.groupId }
    const issues = validateSparkline(next)
    if (issues.length) throw new TypeError(`Invalid sparkline: ${issues.map((issue) => `${issue.path} ${issue.message}`).join('; ')}`)
    this.assertTargetAvailable(next.target, id)
    if (current.groupId) {
      const peers = this.groups.get(current.groupId)?.memberIds.filter((memberId) => memberId !== id) ?? []
      if (peers.some((memberId) => this.require(memberId).type !== next.type)) throw new Error('Grouped sparklines must have the same type')
    }
    this.sparklines.set(id, next)
    this.emit()
    return clone(next)
  }

  remove(id: string): boolean {
    const spec = this.sparklines.get(id)
    if (!spec) return false
    this.sparklines.delete(id)
    if (spec.groupId) this.removeFromGroup(id, spec.groupId)
    this.emit()
    return true
  }

  get(id: string): SparklineSpec | undefined {
    const spec = this.sparklines.get(id)
    return spec ? clone(spec) : undefined
  }

  list(): SparklineSpec[] {
    return [...this.sparklines.values()].map(clone)
  }

  listGroups(): SparklineGroup[] {
    return [...this.groups.values()].map(clone)
  }

  group(memberIds: readonly string[], id = this.idFactory('group')): SparklineGroup {
    const unique = [...new Set(memberIds)]
    if (unique.length < 2) throw new Error('A sparkline group needs at least two members')
    if (!id.trim() || this.groups.has(id)) throw new Error(`Sparkline group ${id || '(empty)'} already exists or is invalid`)
    const members = unique.map((memberId) => this.require(memberId))
    const type = members[0].type
    if (members.some((member) => member.type !== type)) throw new Error('Grouped sparklines must have the same type')
    for (const member of members) if (member.groupId) this.removeFromGroup(member.id, member.groupId)
    const group = { id, memberIds: unique }
    this.groups.set(id, group)
    for (const member of members) this.sparklines.set(member.id, { ...member, groupId: id })
    this.emit()
    return clone(group)
  }

  /** Removes selected members from a group. A group with fewer than two
   * members is dissolved so snapshots never contain degenerate groups. */
  ungroup(memberIds: readonly string[]): void {
    for (const memberId of new Set(memberIds)) {
      const spec = this.sparklines.get(memberId)
      if (spec?.groupId) this.removeFromGroup(memberId, spec.groupId)
    }
    this.emit()
  }

  render(id: string, read: SparklineValueReader, viewport: SparklineViewport): SparklineGeometry {
    const spec = this.require(id)
    const values = this.values(spec, read)
    let domain: Partial<{ min: number; max: number }> | undefined
    if (spec.groupId) {
      const group = this.groups.get(spec.groupId)
      const extents = group?.memberIds
        .map((memberId) => numericExtent(this.values(this.require(memberId), read)))
        .filter((extent): extent is { min: number; max: number } => extent !== null) ?? []
      if (extents.length) domain = {
        min: Math.min(...extents.map((extent) => extent.min)),
        max: Math.max(...extents.map((extent) => extent.max)),
      }
    }
    return compileSparklineGeometry({ type: spec.type, values, viewport, options: spec.options, domain })
  }

  serialize(): SparklineSnapshotV1 {
    return { version: 1, sparklines: this.list(), groups: this.listGroups() }
  }

  /** Replaces manager state atomically; malformed snapshots leave current
   * state untouched. */
  hydrate(snapshot: SparklineSnapshotV1): void {
    const candidate = clone(snapshot)
    const issues = validateSparklineSnapshot(candidate)
    if (issues.length) throw new TypeError(`Invalid sparkline snapshot: ${issues.map((issue) => `${issue.path} ${issue.message}`).join('; ')}`)
    this.sparklines.clear()
    this.groups.clear()
    for (const spec of candidate.sparklines) this.sparklines.set(spec.id, spec)
    for (const group of candidate.groups) this.groups.set(group.id, group)
    this.emit()
  }

  private values(spec: SparklineSpec, read: SparklineValueReader): Array<number | null> {
    return extractSparklineValues(read(spec.source), spec.source, spec.options?.emptyCells)
  }

  private require(id: string): SparklineSpec {
    const spec = this.sparklines.get(id)
    if (!spec) throw new Error(`Unknown sparkline ${id}`)
    return spec
  }

  private assertTargetAvailable(target: SparklineSpec['target'], exceptId?: string): void {
    for (const spec of this.sparklines.values()) {
      if (spec.id !== exceptId && spec.target.sheetId === target.sheetId && spec.target.row === target.row && spec.target.column === target.column) {
        throw new Error(`Target cell already contains sparkline ${spec.id}`)
      }
    }
  }

  private removeFromGroup(memberId: string, groupId: string): void {
    const member = this.sparklines.get(memberId)
    if (member) this.sparklines.set(memberId, { ...member, groupId: undefined })
    const group = this.groups.get(groupId)
    if (!group) return
    const remaining = group.memberIds.filter((id) => id !== memberId)
    if (remaining.length < 2) {
      this.groups.delete(groupId)
      for (const id of remaining) {
        const spec = this.sparklines.get(id)
        if (spec) this.sparklines.set(id, { ...spec, groupId: undefined })
      }
    } else {
      this.groups.set(groupId, { ...group, memberIds: remaining })
    }
  }

  private emit(): void {
    if (!this.listeners.size) return
    const snapshot = this.serialize()
    for (const listener of this.listeners) listener(snapshot)
  }
}
