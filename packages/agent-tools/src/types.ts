export type JsonPrimitive = string | number | boolean | null
export type JsonValue = JsonPrimitive | JsonObject | JsonValue[]
export interface JsonObject { [key: string]: JsonValue }

/** A dependency-free JSON Schema subset. Adapters may use additional JSON Schema keywords as plain JSON. */
export type AgentInputSchema = JsonObject & Readonly<{
  type: 'object'
  properties?: Readonly<Record<string, JsonObject>>
  required?: readonly string[]
  additionalProperties?: boolean | JsonObject
}>

export const AGENT_CHANGESET_PROTOCOL = 'injoffice.agent-changeset' as const
export const AGENT_CHANGESET_PROTOCOL_VERSION = 1 as const
export const AGENT_TOOLS_PROTOCOL = 'injoffice.agent-tools' as const
export const AGENT_TOOLS_PROTOCOL_VERSION = 1 as const

export interface AgentActor {
  /** Stable host identity; never use a display name for authorization. */
  id: string
  kind: 'agent'
  displayName?: string
  provider?: string
  model?: string
  metadata?: JsonObject
}

export interface AgentArtifactIdentity {
  artifactId: string
  format: string
  mediaType?: string
  /** Opaque authoritative revision, compared exactly. */
  revision: string
  /** Content or semantic fingerprint supplied by the adapter. */
  fingerprint: string
}

export interface AgentArtifactEvidence {
  kind: string
  description?: string
  data?: JsonValue
}

export interface AgentIssue {
  severity: 'error' | 'warning' | 'refusal'
  code: string
  message: string
  path?: string
  operationId?: string
  retryable?: boolean
  details?: JsonObject
}

export interface AgentCapability {
  /** Stable operation name, for example `sheet.cell.set`. */
  name: string
  description: string
  inputSchema: AgentInputSchema
  destructive?: boolean
  requiresConfirmation?: boolean
  metadata?: JsonObject
}

export interface AgentCapabilitiesResult {
  protocol: typeof AGENT_TOOLS_PROTOCOL
  protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION
  identity: AgentArtifactIdentity
  capabilities: readonly AgentCapability[]
}

export interface AgentBoundedRequest {
  query?: JsonObject
  cursor?: string
  maxItems?: number
  maxBytes?: number
  signal?: AbortSignal
}

export interface AgentBoundedAdapterResult {
  data: JsonValue
  itemCount: number
  truncated: boolean
  nextCursor?: string
  issues?: readonly AgentIssue[]
  evidence?: readonly AgentArtifactEvidence[]
}

export interface AgentBoundedResult extends AgentBoundedAdapterResult {
  protocol: typeof AGENT_TOOLS_PROTOCOL
  protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION
  identity: AgentArtifactIdentity
  appliedLimits: { maxItems: number; maxBytes: number }
}

export interface AgentOperationInput {
  name: string
  input: JsonObject
  operationId?: string
}

export interface AgentOperation {
  operationId: string
  name: string
  input: JsonObject
}

export interface AgentChangeSetEnvelope {
  protocol: typeof AGENT_CHANGESET_PROTOCOL
  protocolVersion: typeof AGENT_CHANGESET_PROTOCOL_VERSION
  changeSetId: string
  artifactId: string
  format: string
  baseRevision: string
  baseFingerprint: string
  actor: AgentActor
  operations: readonly AgentOperation[]
  metadata?: JsonObject
}

export interface AgentPlanOptions {
  expectedRevision?: string
  expectedFingerprint?: string
  metadata?: JsonObject
  signal?: AbortSignal
}

export interface AgentValidationReport {
  protocol: typeof AGENT_TOOLS_PROTOCOL
  protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION
  changeSetId: string
  valid: boolean
  issues: readonly AgentIssue[]
}

export type AgentValidationResult = AgentValidationReport

export interface AgentArtifactView {
  protocol: typeof AGENT_TOOLS_PROTOCOL
  protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION
  changeSetId: string
  identity: AgentArtifactIdentity
  data: JsonValue
  issues: readonly AgentIssue[]
  evidence: readonly AgentArtifactEvidence[]
}

export interface AgentVerificationCheck {
  name: string
  passed: boolean
  message?: string
  evidence?: readonly AgentArtifactEvidence[]
}

export interface AgentVerificationResult {
  protocol: typeof AGENT_TOOLS_PROTOCOL
  protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION
  changeSetId: string
  verified: boolean
  checks: readonly AgentVerificationCheck[]
  issues: readonly AgentIssue[]
}

export interface AgentAdapterCommitResult<TArtifact = unknown> {
  /** Replacement artifact. May be the same mutable host object or a new immutable value. */
  artifact: TArtifact
  identity: AgentArtifactIdentity
  data?: JsonValue
  evidence?: readonly AgentArtifactEvidence[]
  /** True only when the adapter durably deduplicated the same key. */
  deduplicated?: boolean
}

export interface AgentCommitResult {
  protocol: typeof AGENT_TOOLS_PROTOCOL
  protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION
  changeSetId: string
  idempotencyKey: string
  identity: AgentArtifactIdentity
  data?: JsonValue
  evidence?: readonly AgentArtifactEvidence[]
  deduplicated?: boolean
  /** Verification performed against the exact committed replacement artifact. */
  verification: AgentVerificationResult
}

export interface AgentCommitOptions {
  idempotencyKey: string
  confirmation?: JsonValue
  signal?: AbortSignal
}

export interface AgentRestoreRequest {
  versionId: string
  expectedRevision: string
  expectedFingerprint?: string
  idempotencyKey: string
  confirmation?: JsonValue
  signal?: AbortSignal
}

export interface AgentAdapterIdentityContext<TArtifact> {
  artifact: TArtifact
  actor: Readonly<AgentActor>
}

export interface AgentAdapterContext<TArtifact> extends AgentAdapterIdentityContext<TArtifact> {
  sourceIdentity: Readonly<AgentArtifactIdentity>
  limits: Readonly<AgentSessionLimits>
}

export interface AgentAdapterOperationRequest<TArtifact> extends AgentAdapterContext<TArtifact> {
  changeSet: Readonly<AgentChangeSetEnvelope>
  signal?: AbortSignal
}

export interface AgentAdapterCommitRequest<TArtifact> extends AgentAdapterOperationRequest<TArtifact> {
  idempotencyKey: string
  expectedRevision: string
  expectedFingerprint: string
}

export interface AgentAdapterVerifyRequest<TArtifact> extends AgentAdapterOperationRequest<TArtifact> {
  stage: 'planned' | 'committed'
  /** Present for committed verification and bound to `artifact`/`sourceIdentity`. */
  commit?: Readonly<AgentAdapterCommitResult<TArtifact>>
}

export interface AgentAdapterRestoreRequest<TArtifact> extends AgentAdapterContext<TArtifact> {
  versionId: string
  idempotencyKey: string
  expectedRevision: string
  expectedFingerprint?: string
  signal?: AbortSignal
}

/**
 * Format boundary implemented by DOCX/XLSX/PPTX/PDF integrations. Commit must
 * atomically enforce both expected identity fields and durably deduplicate its
 * idempotency key when the artifact lives outside the current process.
 */
export interface AgentArtifactAdapter<TArtifact = unknown> {
  id: string
  format: string
  supports(artifact: unknown): artifact is TArtifact
  identity(context: AgentAdapterIdentityContext<TArtifact> & { signal?: AbortSignal }): Promise<AgentArtifactIdentity>
  capabilities(context: AgentAdapterContext<TArtifact> & { signal?: AbortSignal }): Promise<readonly AgentCapability[]>
  inspect(context: AgentAdapterContext<TArtifact> & Required<Pick<AgentBoundedRequest, 'maxItems' | 'maxBytes'>> & Omit<AgentBoundedRequest, 'maxItems' | 'maxBytes'>): Promise<AgentBoundedAdapterResult>
  read(context: AgentAdapterContext<TArtifact> & Required<Pick<AgentBoundedRequest, 'maxItems' | 'maxBytes'>> & Omit<AgentBoundedRequest, 'maxItems' | 'maxBytes'>): Promise<AgentBoundedAdapterResult>
  validate(request: AgentAdapterOperationRequest<TArtifact>): Promise<readonly AgentIssue[]>
  preview(request: AgentAdapterOperationRequest<TArtifact>): Promise<Omit<AgentArtifactView, 'protocol' | 'protocolVersion' | 'changeSetId' | 'identity'>>
  diff(request: AgentAdapterOperationRequest<TArtifact>): Promise<Omit<AgentArtifactView, 'protocol' | 'protocolVersion' | 'changeSetId' | 'identity'>>
  verify(request: AgentAdapterVerifyRequest<TArtifact>): Promise<Omit<AgentVerificationResult, 'protocol' | 'protocolVersion' | 'changeSetId'>>
  commit(request: AgentAdapterCommitRequest<TArtifact>): Promise<AgentAdapterCommitResult<TArtifact>>
  restore?(request: AgentAdapterRestoreRequest<TArtifact>): Promise<AgentAdapterCommitResult<TArtifact>>
}

export interface AgentSessionLimits {
  maxReadItems: number
  maxReadBytes: number
  maxOperations: number
  maxOperationBytes: number
}

export interface AgentDestructiveConfirmationRequest {
  action: 'commit' | 'restore'
  actor: Readonly<AgentActor>
  identity: Readonly<AgentArtifactIdentity>
  changeSet?: Readonly<AgentChangeSetEnvelope>
  destructiveOperations: readonly AgentOperation[]
  confirmation?: JsonValue
}

export type AgentDestructiveConfirmationHook = (request: AgentDestructiveConfirmationRequest) => boolean | Promise<boolean>

export interface CreateAgentSessionOptions<TArtifact> {
  artifact: TArtifact
  actor: AgentActor
  adapter?: AgentArtifactAdapter<TArtifact>
  registry?: { resolve(artifact: TArtifact, format?: string): AgentArtifactAdapter<TArtifact> }
  format?: string
  limits?: Partial<AgentSessionLimits>
  confirmDestructive?: AgentDestructiveConfirmationHook
  signal?: AbortSignal
}

export type AgentToolMethod = 'office.capabilities' | 'office.inspect' | 'office.read' | 'office.plan' | 'office.validate' | 'office.preview' | 'office.diff' | 'office.commit' | 'office.verify' | 'office.restore'

export interface AgentToolDescriptor {
  name: AgentToolMethod
  description: string
  inputSchema: AgentInputSchema
}

export interface AgentToolCall {
  protocol: typeof AGENT_TOOLS_PROTOCOL
  protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION
  requestId: string
  method: AgentToolMethod
  params: JsonObject
}

export type AgentToolCallResult =
  | { protocol: typeof AGENT_TOOLS_PROTOCOL; protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION; requestId: string; ok: true; result: JsonValue }
  | { protocol: typeof AGENT_TOOLS_PROTOCOL; protocolVersion: typeof AGENT_TOOLS_PROTOCOL_VERSION; requestId: string; ok: false; error: AgentErrorData }

export type AgentErrorCode = 'ABORTED' | 'ADAPTER_ERROR' | 'CONFIRMATION_DENIED' | 'CONFIRMATION_REQUIRED' | 'IDEMPOTENCY_CONFLICT' | 'INVALID_ADAPTER_RESULT' | 'INVALID_ARGUMENT' | 'LIMIT_EXCEEDED' | 'RESTORE_UNSUPPORTED' | 'STALE_REVISION' | 'VALIDATION_FAILED'

export interface AgentErrorData {
  code: AgentErrorCode
  message: string
  retryable: boolean
  issues?: readonly AgentIssue[]
  details?: JsonObject
}
