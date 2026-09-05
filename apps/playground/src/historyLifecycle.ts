import { HistoryManager } from '../../../packages/history/src/manager'
import { historyVersionMatches } from '../../../packages/history/src/validation'
import type {
  HistoryAuthor,
  HistoryCreateRequest,
  HistoryHost,
  HistoryListPage,
  HistoryListRequest,
  HistoryLoadRequest,
  HistoryLoadedVersion,
  HistoryVersionInfo,
} from '../../../packages/history/src/types'

export type HistorySnapshot = string

const ARTIFACT_ID = 'playground-history'
const USER: HistoryAuthor = { id: 'playground-user', kind: 'user', displayName: 'Playground user' }
const AGENT: HistoryAuthor = { id: 'playground-agent', kind: 'agent', displayName: 'Playground agent' }

type Stored = { version: HistoryVersionInfo; snapshot: HistorySnapshot }

/** In-memory host: filters before pagination and never mutates stored snapshots. */
export class MemoryHistoryHost implements HistoryHost<HistorySnapshot> {
  private readonly versions: Stored[] = []
  private clock = 1_700_000_000_000
  private sequence = 0

  constructor(private readonly artifactId = ARTIFACT_ID) {}

  seed(snapshot: HistorySnapshot, author: HistoryAuthor, reason: HistoryVersionInfo['reason'], contentType: string, description?: string): HistoryVersionInfo {
    return this.push({ snapshot, author, reason, contentType, description })
  }

  async listVersions(request: Readonly<HistoryListRequest>): Promise<HistoryListPage> {
    const filtered = this.newestFirst().filter((entry) => historyVersionMatches(entry.version, request.filter))
    const start = request.cursor ? filtered.findIndex((entry) => entry.version.id === request.cursor) + 1 : 0
    const limit = request.limit ?? 50
    const slice = start < 1 && request.cursor ? [] : filtered.slice(Math.max(0, start), Math.max(0, start) + limit)
    const next = filtered[Math.max(0, start) + limit]
    return {
      versions: slice.map((entry) => structuredClone(entry.version)),
      ...(next ? { nextCursor: next.version.id } : {}),
    }
  }

  async loadVersion(request: Readonly<HistoryLoadRequest>): Promise<HistoryLoadedVersion<HistorySnapshot>> {
    const found = this.versions.find((entry) => entry.version.id === request.versionId)
    if (!found) throw new Error(`version ${request.versionId} does not exist`)
    return { version: structuredClone(found.version), snapshot: structuredClone(found.snapshot) }
  }

  async createVersion(request: Readonly<HistoryCreateRequest<HistorySnapshot>>): Promise<HistoryVersionInfo> {
    const head = this.newestFirst()[0]
    if (request.expectedHeadVersionId && head?.version.id !== request.expectedHeadVersionId) {
      throw new Error(`expected head ${request.expectedHeadVersionId}, found ${head?.version.id ?? 'none'}`)
    }
    return this.push({
      snapshot: request.snapshot,
      author: request.author,
      reason: request.reason,
      contentType: request.contentType,
      description: request.description,
      sourceVersionId: request.sourceVersionId,
      retention: request.retention,
    })
  }

  private newestFirst(): Stored[] {
    return [...this.versions].sort((left, right) => right.version.sequence - left.version.sequence)
  }

  private push(input: {
    snapshot: HistorySnapshot
    author: HistoryAuthor
    reason: HistoryVersionInfo['reason']
    contentType: string
    description?: string
    sourceVersionId?: string
    retention?: Partial<HistoryVersionInfo['retention']>
  }): HistoryVersionInfo {
    this.sequence += 1
    this.clock += 1_000
    const version: HistoryVersionInfo = {
      id: `v${this.sequence}`,
      artifactId: this.artifactId,
      sequence: this.sequence,
      createdAt: this.clock,
      size: input.snapshot.length,
      contentType: input.contentType,
      author: structuredClone(input.author),
      reason: input.reason,
      retention: {
        policyId: input.retention?.policyId ?? 'playground-session',
        expiresAt: input.retention?.expiresAt === undefined ? this.clock + 86_400_000 : input.retention.expiresAt,
        legalHold: input.retention?.legalHold ?? false,
      },
      ...(input.description === undefined ? {} : { description: input.description }),
      ...(input.sourceVersionId === undefined ? {} : { sourceVersionId: input.sourceVersionId }),
    }
    this.versions.push({ version, snapshot: structuredClone(input.snapshot) })
    return structuredClone(version)
  }
}

export function playgroundHistoryAuthors() {
  return { user: USER, agent: AGENT, artifactId: ARTIFACT_ID }
}

export function createPlaygroundHistory(initial: { snapshot: HistorySnapshot; contentType: string; description: string }[]): {
  host: MemoryHistoryHost
  manager: HistoryManager<HistorySnapshot>
} {
  const host = new MemoryHistoryHost(ARTIFACT_ID)
  for (const entry of initial) host.seed(entry.snapshot, USER, 'save', entry.contentType, entry.description)
  return { host, manager: new HistoryManager(host, ARTIFACT_ID) }
}
