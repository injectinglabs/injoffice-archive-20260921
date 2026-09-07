import {
  createAgentSession,
  type AgentArtifactAdapter,
  type AgentArtifactEvidence,
  type AgentArtifactIdentity,
  type AgentCapability as CoreAgentCapability,
  type AgentIssue,
  type AgentOperation,
  type AgentVerificationResult,
  type JsonObject,
  type JsonValue,
} from '@injoffice/agent-tools'
import {
  createAgentDemoScenario,
  proposeAgentDemoOperations,
  type AgentDemoArtifact,
  type AgentDemoFormat,
  type AgentDemoMode,
  type AgentDemoOperation,
} from './agentDemoScenario'

export type AgentCapability = { operation: string; access: 'read' | 'write'; destructive?: boolean }
export type AgentInspection = { artifactId: string; revision: string; fingerprint: string; selection: string; content: Record<string, unknown>; truncated: boolean }
export type AgentValidation = { ok: boolean; issues: Array<{ code: string; path: string; message: string }> }
export type AgentPreview = { artifact: AgentDemoArtifact; summary: string; evidence: string[] }
export type AgentDiff = { changes: Array<{ target: string; before: string; after: string }> }
export type AgentCommitReceipt = {
  artifactId: string
  previousRevision: string
  revision: string
  fingerprint: string
  operationIds: string[]
  evidence: string[]
  verification: AgentVerificationResult
}
export type AgentVerification = { ok: boolean; revision: string; fingerprint: string; evidence: string[] }

export interface AgentDemoChangeSet {
  id: string
  expectedRevision: string
  operations: readonly AgentDemoOperation[]
  preview(): Promise<AgentPreview>
  diff(): Promise<AgentDiff>
  validate(): Promise<AgentValidation>
  commit(input: { expectedRevision: string; idempotencyKey: string; confirmation: 'approved' }): Promise<AgentCommitReceipt>
  verify(receipt: AgentCommitReceipt): Promise<AgentVerification>
}

export interface AgentDemoSession {
  capabilities(): Promise<{ format: AgentDemoFormat; operations: AgentCapability[] }>
  inspect(input: { selection: string; maxItems: number; maxBytes: number }): Promise<AgentInspection>
  plan(operations: readonly AgentDemoOperation[], input: { expectedRevision: string }): Promise<AgentDemoChangeSet>
}

const WRITE_OPERATIONS: Record<AgentDemoFormat, Array<{ name: string; description: string; destructive?: boolean }>> = {
  xlsx: [
    { name: 'xlsx.cell.set_value', description: 'Set one cell value using a stable sheet id and zero-based cell coordinates.' },
    { name: 'xlsx.cell.set_formula', description: 'Set one cell formula using the guarded native mutation protocol.' },
  ],
  docx: [{ name: 'docx.text.replace', description: 'Replace one guarded native text target.' }],
  pptx: [{ name: 'pptx.authored.slide.update', description: 'Update one stable slide in an authored deck.' }],
  pdf: [
    { name: 'pdf.page.rotate', description: 'Rotate one or more bounded PDF pages.' },
    { name: 'pdf.page.delete', description: 'Delete bounded pages from the source PDF.', destructive: true },
  ],
}

const MEDIA_TYPES: Record<AgentDemoFormat, string> = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf: 'application/pdf',
}

const EMPTY_OBJECT_SCHEMA = { type: 'object', properties: {}, additionalProperties: true } as const

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value)
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`
  return `{${Object.entries(value as Record<string, unknown>).sort(([a], [b]) => a.localeCompare(b)).map(([key, item]) => `${JSON.stringify(key)}:${stableStringify(item)}`).join(',')}}`
}

function fingerprint(value: unknown): string {
  let hash = 0x811c9dc5
  for (const char of stableStringify(value)) {
    hash ^= char.charCodeAt(0)
    hash = Math.imul(hash, 0x01000193)
  }
  return `fnv1a32:${(hash >>> 0).toString(16).padStart(8, '0')}`
}

function artifactIdentity(artifact: AgentDemoArtifact): AgentArtifactIdentity {
  return { artifactId: artifact.artifactId, format: artifact.format, mediaType: MEDIA_TYPES[artifact.format], revision: artifact.revision, fingerprint: fingerprint(artifact.content) }
}

function operationInput(operation: AgentDemoOperation): JsonObject {
  return operation.input as JsonObject
}

function demoOperation(operation: AgentOperation): AgentDemoOperation {
  if (operation.name === 'xlsx.cell.set_value') {
    const cell = operation.input.cell as JsonObject
    return { operationId: operation.operationId, op: operation.name, target: `${String(operation.input.sheetId)}!D${Number(cell.row) + 1}`, value: operation.input.value, input: operation.input }
  }
  if (operation.name === 'docx.text.replace') return { operationId: operation.operationId, op: operation.name, target: String(operation.input.targetId), value: operation.input.text, input: operation.input }
  if (operation.name === 'pptx.authored.slide.update') {
    const patch = operation.input.patch as JsonObject
    const metric = patch.metric as JsonObject
    return { operationId: operation.operationId, op: operation.name, target: `${String(operation.input.slideId)}/shape-readiness`, value: metric.value, input: operation.input }
  }
  if (operation.name === 'pdf.page.rotate') {
    const pages = operation.input.pages as JsonValue[]
    return { operationId: operation.operationId, op: operation.name, target: `page-${String(pages[0])}`, value: operation.input.degrees, input: operation.input }
  }
  return { operationId: operation.operationId, op: operation.name, target: '', input: operation.input }
}

function beforeValue(artifact: AgentDemoArtifact, operation: AgentDemoOperation): unknown {
  if (artifact.format === 'xlsx') return (artifact.content.rows as unknown[][])[3]?.[3]
  if (artifact.format === 'docx') return (artifact.content.blocks as Array<{ id: string; text: string }>).find((block) => block.id === operation.target)?.text
  if (artifact.format === 'pptx') return (artifact.content.metric as { value: string }).value
  return artifact.content.rotation
}

function applyOperation(artifact: AgentDemoArtifact, operation: AgentDemoOperation): AgentDemoArtifact {
  const next = structuredClone(artifact)
  if (operation.op === 'xlsx.cell.set_value') (next.content.rows as unknown[][])[3][3] = operation.value
  if (operation.op === 'docx.text.replace') {
    const block = (next.content.blocks as Array<{ id: string; text: string }>).find((item) => item.id === operation.target)
    if (block) block.text = String(operation.value)
  }
  if (operation.op === 'pptx.authored.slide.update') (next.content.metric as { value: string }).value = String(operation.value)
  if (operation.op === 'pdf.page.rotate') next.content.rotation = Number(operation.value)
  return next
}

function supported(format: AgentDemoFormat, operation: AgentOperation): boolean {
  return WRITE_OPERATIONS[format].some((capability) => capability.name === operation.name)
}

function refusal(format: AgentDemoFormat, operation: AgentOperation): AgentIssue {
  return {
    severity: 'refusal', code: 'UNSUPPORTED_OPERATION', path: `/operations/${operation.operationId}/name`, operationId: operation.operationId,
    message: `${operation.name} is not supported by the ${format.toUpperCase()} adapter; no output was written.`,
  }
}

function evidence(kind: string, description: string): AgentArtifactEvidence { return { kind, description } }

/** The adapter is demo-specific; session safety and lifecycle are the shipped package. */
function createDocumentShapedDemoAdapter(format: AgentDemoFormat): AgentArtifactAdapter<AgentDemoArtifact> {
  const capabilities: CoreAgentCapability[] = [
    { name: `${format}.inspect`, description: `Inspect bounded ${format.toUpperCase()} document data.`, inputSchema: EMPTY_OBJECT_SCHEMA },
    ...WRITE_OPERATIONS[format].map((item) => ({
      name: item.name, description: item.description, inputSchema: EMPTY_OBJECT_SCHEMA,
      ...(item.destructive ? { destructive: true, requiresConfirmation: true } : {}),
    })),
  ]
  return {
    id: `playground-${format}`,
    format,
    supports(artifact): artifact is AgentDemoArtifact {
      return Boolean(artifact && typeof artifact === 'object' && (artifact as AgentDemoArtifact).format === format)
    },
    async identity({ artifact }) { return artifactIdentity(artifact) },
    async capabilities() { return capabilities },
    async inspect({ artifact }) {
      return { data: artifact.content as JsonValue, itemCount: 1, truncated: false, evidence: [evidence('bounded-inspection', `Read ${artifact.name} at ${artifact.revision}.`)] }
    },
    async read({ artifact }) {
      return { data: artifact.content as JsonValue, itemCount: 1, truncated: false, evidence: [evidence('bounded-read', `Read one document projection from ${artifact.revision}.`)] }
    },
    async validate({ artifact, changeSet }) {
      return changeSet.operations.filter((operation) => !supported(artifact.format, operation)).map((operation) => refusal(artifact.format, operation))
    },
    async preview({ artifact, changeSet }) {
      const invalid = changeSet.operations.filter((operation) => !supported(artifact.format, operation))
      let preview = structuredClone(artifact)
      if (invalid.length === 0) for (const operation of changeSet.operations) preview = applyOperation(preview, demoOperation(operation))
      return { data: preview as unknown as JsonValue, issues: invalid.map((operation) => refusal(artifact.format, operation)), evidence: [evidence('isolated-preview', 'Preview was produced without replacing the source artifact.')] }
    },
    async diff({ artifact, changeSet }) {
      const changes = changeSet.operations.filter((operation) => supported(artifact.format, operation)).map((operation) => {
        const demo = demoOperation(operation)
        return { target: demo.target, before: String(beforeValue(artifact, demo) ?? 'blank'), after: String(demo.value ?? 'blank') }
      })
      return { data: { changes }, issues: [], evidence: [evidence('semantic-diff', `${changes.length} bounded target changed.`)] }
    },
    async verify({ artifact, changeSet, stage, commit }) {
      const invalid = changeSet.operations.filter((operation) => !supported(artifact.format, operation))
      const identityMatches = stage === 'planned' || Boolean(commit && commit.identity.fingerprint === artifactIdentity(artifact).fingerprint)
      const checks = [
        { name: 'all-operations-supported', passed: invalid.length === 0, message: invalid.length ? 'One or more operations are unsupported.' : 'Every operation is advertised by the adapter.' },
        { name: stage === 'planned' ? 'preview-isolated' : 'replacement-reopened', passed: identityMatches, message: stage === 'planned' ? 'The source artifact remains authoritative.' : 'The replacement fingerprint matches the commit receipt.' },
      ]
      return { verified: checks.every((check) => check.passed), checks, issues: invalid.map((operation) => refusal(artifact.format, operation)) }
    },
    async commit({ artifact, changeSet, expectedRevision, expectedFingerprint }) {
      const identity = artifactIdentity(artifact)
      if (identity.revision !== expectedRevision || identity.fingerprint !== expectedFingerprint) throw new Error('Artifact changed before the atomic write.')
      let next = structuredClone(artifact)
      for (const operation of changeSet.operations) next = applyOperation(next, demoOperation(operation))
      next.revision = `rev-${String(Number(artifact.revision.replace(/\D/g, '')) + 1).padStart(3, '0')}`
      return {
        artifact: next,
        identity: artifactIdentity(next),
        data: next.content as JsonValue,
        evidence: [evidence('expected-identity', 'Expected revision and fingerprint matched.'), evidence('atomic-commit', 'The bounded operation list was applied as one replacement.')],
      }
    },
  }
}

/**
 * Tiny UI facade over @injoffice/agent-tools. Only the model proposal is
 * deterministic simulation; createAgentSession owns limits, immutable plans,
 * revision/fingerprint guards, confirmation, idempotency, and verification.
 */
function createAgentDemoSession(artifact: AgentDemoArtifact): AgentDemoSession {
  const coreSession = createAgentSession({
    artifact,
    adapter: createDocumentShapedDemoAdapter(artifact.format),
    actor: { id: 'agent-local-analyst', kind: 'agent', displayName: 'Local analyst' },
    limits: { maxReadItems: 20, maxReadBytes: 16_000, maxOperations: 8, maxOperationBytes: 8_000 },
    confirmDestructive: async ({ confirmation }) => confirmation === 'approved',
  })
  return {
    async capabilities() {
      const report = await (await coreSession).capabilities()
      return {
        format: artifact.format,
        operations: report.capabilities.map((capability) => ({
          operation: capability.name,
          access: capability.name.endsWith('.inspect') ? 'read' : 'write',
          ...(capability.destructive ? { destructive: true } : {}),
        })),
      }
    },
    async inspect({ selection, maxItems, maxBytes }) {
      const result = await (await coreSession).inspect({ query: { selection }, maxItems, maxBytes })
      return {
        artifactId: result.identity.artifactId, revision: result.identity.revision, fingerprint: result.identity.fingerprint,
        selection, content: result.data as Record<string, unknown>, truncated: result.truncated,
      }
    },
    async plan(operations, { expectedRevision }) {
      const session = await coreSession
      const changeSet = await session.plan(operations.map((operation) => ({ name: operation.op, input: operationInput(operation), operationId: operation.operationId })), {
        expectedRevision,
        expectedFingerprint: session.identity.fingerprint,
      })
      return {
        id: changeSet.envelope.changeSetId,
        expectedRevision: changeSet.envelope.baseRevision,
        operations: structuredClone(operations),
        async validate() {
          const result = await changeSet.validate()
          return { ok: result.valid, issues: result.issues.map((issue) => ({ code: issue.code.toLowerCase(), path: issue.path ?? `/operations/${issue.operationId ?? ''}`, message: issue.message })) }
        },
        async preview() {
          const result = await changeSet.preview()
          return { artifact: result.data as unknown as AgentDemoArtifact, summary: `${operations.length} bounded operation${operations.length === 1 ? '' : 's'} against ${expectedRevision}`, evidence: result.evidence.map((item) => item.description ?? item.kind) }
        },
        async diff() {
          const result = await changeSet.diff()
          return result.data as unknown as AgentDiff
        },
        async commit(input) {
          if (input.expectedRevision !== changeSet.envelope.baseRevision) throw new Error('Commit revision does not match the reviewed change set.')
          const result = await changeSet.commit({ idempotencyKey: input.idempotencyKey, confirmation: input.confirmation })
          return {
            artifactId: result.identity.artifactId,
            previousRevision: changeSet.envelope.baseRevision,
            revision: result.identity.revision,
            fingerprint: result.identity.fingerprint,
            operationIds: changeSet.envelope.operations.map((operation) => operation.operationId),
            evidence: result.evidence?.map((item) => item.description ?? item.kind) ?? [],
            verification: result.verification,
          }
        },
        async verify(receipt) {
          return { ok: receipt.verification.verified, revision: receipt.revision, fingerprint: receipt.fingerprint, evidence: receipt.verification.checks.map((check) => check.message ?? check.name) }
        },
      }
    },
  }
}

export function createDemoSessionInput(format: AgentDemoFormat, mode: AgentDemoMode) {
  const scenario = createAgentDemoScenario(format, mode)
  return { scenario, operations: proposeAgentDemoOperations(format, mode), session: createAgentDemoSession(scenario.artifact) }
}
