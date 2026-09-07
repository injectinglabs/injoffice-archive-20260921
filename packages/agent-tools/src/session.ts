import { AgentToolsError, asAgentToolsError } from './errors'
import { canonicalJson, cloneJson, immutableJson, jsonBytes, stableId } from './json'
import type {
  AgentAdapterCommitResult,
  AgentArtifactAdapter,
  AgentArtifactIdentity,
  AgentArtifactView,
  AgentActor,
  AgentBoundedAdapterResult,
  AgentBoundedRequest,
  AgentBoundedResult,
  AgentCapabilitiesResult,
  AgentCapability,
  AgentChangeSetEnvelope,
  AgentCommitOptions,
  AgentCommitResult,
  AgentIssue,
  AgentOperation,
  AgentOperationInput,
  AgentPlanOptions,
  AgentRestoreRequest,
  AgentSessionLimits,
  AgentValidationResult,
  AgentVerificationResult,
  CreateAgentSessionOptions,
  JsonObject,
} from './types'
import { AGENT_CHANGESET_PROTOCOL, AGENT_CHANGESET_PROTOCOL_VERSION, AGENT_TOOLS_PROTOCOL, AGENT_TOOLS_PROTOCOL_VERSION } from './types'

export const DEFAULT_AGENT_SESSION_LIMITS: Readonly<AgentSessionLimits> = Object.freeze({
  maxReadItems: 1_000,
  maxReadBytes: 1_048_576,
  maxOperations: 100,
  maxOperationBytes: 1_048_576,
})

type StoredCommit = { signature: string; promise: Promise<AgentCommitResult>; mutationStarted: boolean }

export class AgentChangeSet<TArtifact> {
  readonly #session: AgentSession<TArtifact>
  constructor(readonly envelope: Readonly<AgentChangeSetEnvelope>, session: AgentSession<TArtifact>) { this.#session = session }

  validate(signal?: AbortSignal): Promise<AgentValidationResult> { return this.#session.validate(this, signal) }
  preview(signal?: AbortSignal): Promise<AgentArtifactView> { return this.#session.preview(this, signal) }
  diff(signal?: AbortSignal): Promise<AgentArtifactView> { return this.#session.diff(this, signal) }
  verify(signal?: AbortSignal): Promise<AgentVerificationResult> { return this.#session.verify(this, signal) }
  commit(options: AgentCommitOptions): Promise<AgentCommitResult> { return this.#session.commit(this, options) }
  toJSON(): Readonly<AgentChangeSetEnvelope> { return this.envelope }
}

export class AgentSession<TArtifact> {
  #artifact: TArtifact
  #identity: Readonly<AgentArtifactIdentity>
  readonly #commits = new Map<string, StoredCommit>()
  readonly #changeSets = new WeakSet<AgentChangeSet<TArtifact>>()
  #mutationTail: Promise<void> = Promise.resolve()
  readonly #confirm?: CreateAgentSessionOptions<TArtifact>['confirmDestructive']
  readonly actor: Readonly<AgentActor>
  readonly adapter: AgentArtifactAdapter<TArtifact>
  readonly limits: Readonly<AgentSessionLimits>

  constructor(
    artifact: TArtifact,
    identity: AgentArtifactIdentity,
    actor: Readonly<AgentActor>,
    adapter: AgentArtifactAdapter<TArtifact>,
    limits: Readonly<AgentSessionLimits>,
    confirm?: CreateAgentSessionOptions<TArtifact>['confirmDestructive'],
  ) {
    const checkedActor = immutableJson(actor, 'actor')
    validateActor(checkedActor)
    const checkedIdentity = adapterIdentity(identity, 'artifact identity')
    const checkedLimits = immutableJson(limits, 'session limits')
    for (const [name, value] of Object.entries(checkedLimits)) if (!Number.isSafeInteger(value) || value < 1) throw new AgentToolsError('INVALID_ARGUMENT', `${name} must be a positive safe integer.`)
    if (checkedIdentity.format !== adapter.format) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Adapter format ${adapter.format} does not match artifact format ${checkedIdentity.format}.`)
    if (!adapter.supports(artifact)) throw new AgentToolsError('INVALID_ARGUMENT', `Adapter ${adapter.id} does not support the artifact.`)
    this.#artifact = artifact
    this.#identity = checkedIdentity
    this.actor = checkedActor
    this.adapter = adapter
    this.limits = checkedLimits
    this.#confirm = confirm
  }

  get artifact(): TArtifact { return this.#artifact }
  get identity(): Readonly<AgentArtifactIdentity> { return this.#identity }

  async capabilities(signal?: AbortSignal): Promise<AgentCapabilitiesResult> {
    const identity = await this.#refreshIdentity(signal)
    let raw: readonly AgentCapability[]
    try { raw = await this.adapter.capabilities({ ...this.#context(identity), signal }) }
    catch (error) { throw asAgentToolsError(error) }
    abortIfNeeded(signal)
    const names = new Set<string>()
    const rawCapabilities = adapterJson(raw, 'adapter capabilities')
    if (!Array.isArray(rawCapabilities)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter capabilities must be an array.')
    const capabilities = rawCapabilities.map((entry, index) => {
      const capability = entry
      if (!record(capability)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Capability ${index} must be an object.`)
      exactAdapterKeys(capability, ['name', 'description', 'inputSchema', 'destructive', 'requiresConfirmation', 'metadata'], `Capability ${index}`)
      if (typeof capability.name !== 'string' || !validName(capability.name) || typeof capability.description !== 'string' || !capability.description.trim()) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Capability ${index} has an invalid name or description.`)
      const name = capability.name
      if (names.has(name)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Capability ${name} is duplicated.`)
      names.add(name)
      if (!record(capability.inputSchema) || capability.inputSchema.type !== 'object') throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Capability ${name} must advertise an object inputSchema.`)
      if ((capability.destructive !== undefined && typeof capability.destructive !== 'boolean') || (capability.requiresConfirmation !== undefined && typeof capability.requiresConfirmation !== 'boolean') || (capability.metadata !== undefined && !record(capability.metadata))) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Capability ${name} has invalid optional fields.`)
      return capability as unknown as AgentCapability
    }).sort((left, right) => left.name.localeCompare(right.name))
    return immutableJson({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, identity, capabilities }, 'capabilities result')
  }

  inspect(request: AgentBoundedRequest = {}): Promise<AgentBoundedResult> { return this.#bounded('inspect', request) }
  read(request: AgentBoundedRequest = {}): Promise<AgentBoundedResult> { return this.#bounded('read', request) }

  async plan(operations: readonly AgentOperationInput[], options: AgentPlanOptions = {}): Promise<AgentChangeSet<TArtifact>> {
    abortIfNeeded(options.signal)
    if (!Array.isArray(operations) || operations.length === 0) throw new AgentToolsError('INVALID_ARGUMENT', 'A change set requires at least one operation.')
    if (operations.length > this.limits.maxOperations) throw new AgentToolsError('LIMIT_EXCEEDED', `A change set may contain at most ${this.limits.maxOperations} operations.`)
    const clonedOperations = cloneJson(operations, 'operations')
    const identity = await this.#refreshIdentity(options.signal)
    assertExpected(identity, options.expectedRevision, options.expectedFingerprint)
    const ids = new Set<string>()
    const planned: AgentOperation[] = clonedOperations.map((raw, index) => {
      const operation = raw
      if (!record(operation)) throw new AgentToolsError('INVALID_ARGUMENT', `Operation ${index} must be an object.`)
      exactInputKeys(operation, ['operationId', 'name', 'input'], `Operation ${index}`)
      if (typeof operation.name !== 'string' || !validName(operation.name)) throw new AgentToolsError('INVALID_ARGUMENT', `Operation ${index} has an invalid name.`)
      const operationId = operation.operationId ?? `op-${String(index + 1).padStart(4, '0')}`
      if (typeof operationId !== 'string' || !validId(operationId)) throw new AgentToolsError('INVALID_ARGUMENT', `Operation ${index} has an invalid operationId.`)
      if (!record(operation.input)) throw new AgentToolsError('INVALID_ARGUMENT', `Operation ${index} input must be a JSON object.`)
      if (ids.has(operationId)) throw new AgentToolsError('INVALID_ARGUMENT', `Operation id ${operationId} is duplicated.`)
      ids.add(operationId)
      return { operationId, name: operation.name, input: operation.input as JsonObject }
    })
    if (jsonBytes(planned) > this.limits.maxOperationBytes) throw new AgentToolsError('LIMIT_EXCEEDED', `Change set operations exceed ${this.limits.maxOperationBytes} JSON bytes.`)
    const content = {
      artifactId: identity.artifactId,
      format: identity.format,
      baseRevision: identity.revision,
      baseFingerprint: identity.fingerprint,
      actor: this.actor,
      operations: planned,
      ...(options.metadata ? { metadata: options.metadata } : {}),
    }
    const envelope: AgentChangeSetEnvelope = {
      protocol: AGENT_CHANGESET_PROTOCOL,
      protocolVersion: AGENT_CHANGESET_PROTOCOL_VERSION,
      changeSetId: `changeset-${stableId(content)}`,
      ...content,
    }
    return this.#track(immutableJson(envelope, 'change set'))
  }

  loadChangeSet(envelope: AgentChangeSetEnvelope): AgentChangeSet<TArtifact> {
    const cloned = immutableJson(envelope, 'change set')
    validateEnvelopeShape(cloned, this.limits)
    if (cloned.protocol !== AGENT_CHANGESET_PROTOCOL || cloned.protocolVersion !== AGENT_CHANGESET_PROTOCOL_VERSION) throw new AgentToolsError('INVALID_ARGUMENT', 'Unsupported change set protocol or version.')
    const { changeSetId: _ignored, protocol: _protocol, protocolVersion: _protocolVersion, ...content } = cloned
    if (cloned.changeSetId !== `changeset-${stableId(content)}`) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set id does not match its canonical contents.')
    if (cloned.artifactId !== this.identity.artifactId || cloned.format !== this.identity.format) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set belongs to another artifact or format.')
    if (canonicalJson(cloned.actor) !== canonicalJson(this.actor)) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set belongs to another actor.')
    return this.#track(cloned)
  }

  async validate(changeSet: AgentChangeSet<TArtifact>, signal?: AbortSignal): Promise<AgentValidationResult> {
    const envelope = this.#owned(changeSet)
    const advertised = await this.capabilities(signal)
    const identity = advertised.identity
    const issues: AgentIssue[] = []
    if (!sameIdentity(identity, baseIdentity(envelope))) issues.push(staleIssue(envelope, identity))
    const available = new Set(advertised.capabilities.map(({ name }) => name))
    for (const operation of envelope.operations) if (!available.has(operation.name)) issues.push({ severity: 'refusal', code: 'UNSUPPORTED_OPERATION', message: `Adapter does not advertise ${operation.name}.`, operationId: operation.operationId })
    try {
      const adapterIssues = await this.adapter.validate({ ...this.#context(identity), changeSet: envelope, signal })
      abortIfNeeded(signal)
      issues.push(...normalizeIssues(adapterIssues))
    }
    catch (error) { throw asAgentToolsError(error) }
    const normalized = normalizeIssues(issues)
    return immutableJson({
      protocol: AGENT_TOOLS_PROTOCOL,
      protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
      changeSetId: envelope.changeSetId,
      valid: !normalized.some(({ severity }) => severity === 'error' || severity === 'refusal'),
      issues: normalized,
    }, 'validation result')
  }

  preview(changeSet: AgentChangeSet<TArtifact>, signal?: AbortSignal): Promise<AgentArtifactView> { return this.#view('preview', changeSet, signal) }
  diff(changeSet: AgentChangeSet<TArtifact>, signal?: AbortSignal): Promise<AgentArtifactView> { return this.#view('diff', changeSet, signal) }

  async verify(changeSet: AgentChangeSet<TArtifact>, signal?: AbortSignal): Promise<AgentVerificationResult> {
    const envelope = this.#owned(changeSet)
    const identity = await this.#requireCurrent(envelope, signal)
    let raw
    try { raw = await this.adapter.verify({ ...this.#context(identity), changeSet: envelope, stage: 'planned', signal }) }
    catch (error) { throw asAgentToolsError(error) }
    abortIfNeeded(signal)
    return verification(envelope.changeSetId, raw)
  }

  commit(changeSet: AgentChangeSet<TArtifact>, options: AgentCommitOptions): Promise<AgentCommitResult> {
    const envelope = this.#owned(changeSet)
    validIdempotencyKey(options.idempotencyKey)
    // Bind idempotency to the full canonical mutation, not the compact
    // non-cryptographic display id, so even an adversarial id collision cannot
    // alias two different writes.
    const signature = `commit:${canonicalJson(envelope)}`
    const stored = this.#commits.get(options.idempotencyKey)
    if (stored) {
      if (stored.signature !== signature) throw new AgentToolsError('IDEMPOTENCY_CONFLICT', 'Idempotency key was already used for another mutation.')
      return stored.promise
    }
    const entry: StoredCommit = { signature, promise: undefined as unknown as Promise<AgentCommitResult>, mutationStarted: false }
    const promise = this.#serialize(() => this.#commit(envelope, options, () => { entry.mutationStarted = true }))
    entry.promise = promise
    this.#commits.set(options.idempotencyKey, entry)
    void promise.catch(() => { if (!entry.mutationStarted && this.#commits.get(options.idempotencyKey) === entry) this.#commits.delete(options.idempotencyKey) })
    return promise
  }

  restore(request: AgentRestoreRequest): Promise<AgentCommitResult> {
    if (!this.adapter.restore) throw new AgentToolsError('RESTORE_UNSUPPORTED', `Adapter ${this.adapter.id} does not support restore.`)
    validIdempotencyKey(request.idempotencyKey)
    if (!request.versionId.trim()) throw new AgentToolsError('INVALID_ARGUMENT', 'versionId is required.')
    const signature = `restore:${request.versionId}:${request.expectedRevision}:${request.expectedFingerprint ?? ''}`
    const stored = this.#commits.get(request.idempotencyKey)
    if (stored) {
      if (stored.signature !== signature) throw new AgentToolsError('IDEMPOTENCY_CONFLICT', 'Idempotency key was already used for another mutation.')
      return stored.promise
    }
    const entry: StoredCommit = { signature, promise: undefined as unknown as Promise<AgentCommitResult>, mutationStarted: false }
    const promise = this.#serialize(() => this.#restore(request, () => { entry.mutationStarted = true }))
    entry.promise = promise
    this.#commits.set(request.idempotencyKey, entry)
    void promise.catch(() => { if (!entry.mutationStarted && this.#commits.get(request.idempotencyKey) === entry) this.#commits.delete(request.idempotencyKey) })
    return promise
  }

  async #bounded(kind: 'inspect' | 'read', request: AgentBoundedRequest): Promise<AgentBoundedResult> {
    abortIfNeeded(request.signal)
    const maxItems = boundedInteger(request.maxItems, this.limits.maxReadItems, `${kind}.maxItems`)
    const maxBytes = boundedInteger(request.maxBytes, this.limits.maxReadBytes, `${kind}.maxBytes`)
    const identity = await this.#refreshIdentity(request.signal)
    let raw: AgentBoundedAdapterResult
    try { raw = await this.adapter[kind]({ ...this.#context(identity), ...request, maxItems, maxBytes }) }
    catch (error) { throw asAgentToolsError(error) }
    abortIfNeeded(request.signal)
    const value = adapterJson(raw, `${kind} result`)
    if (!record(value)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result must be an object.`)
    exactAdapterKeys(value, ['data', 'itemCount', 'truncated', 'nextCursor', 'issues', 'evidence'], `${kind} result`)
    if (!Object.prototype.hasOwnProperty.call(value, 'data')) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result must include data.`)
    if (typeof value.itemCount !== 'number' || !Number.isSafeInteger(value.itemCount) || value.itemCount < 0 || value.itemCount > maxItems) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result violates its item bound.`)
    const itemCount = value.itemCount
    if (typeof value.truncated !== 'boolean') throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result must include a boolean truncated field.`)
    if (jsonBytes(value.data) > maxBytes) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result violates its byte bound.`)
    if (value.nextCursor !== undefined && typeof value.nextCursor !== 'string') throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result has an invalid cursor.`)
    if (value.nextCursor !== undefined && !value.truncated) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result has a cursor but is not truncated.`)
    const issues = value.issues === undefined ? undefined : normalizeIssues(value.issues)
    const evidence = value.evidence === undefined ? undefined : normalizeEvidence(value.evidence)
    return immutableJson({
      protocol: AGENT_TOOLS_PROTOCOL,
      protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
      identity,
      appliedLimits: { maxItems, maxBytes },
      data: value.data,
      itemCount,
      truncated: value.truncated,
      ...(value.nextCursor !== undefined ? { nextCursor: value.nextCursor } : {}),
      ...(issues ? { issues } : {}),
      ...(evidence ? { evidence } : {}),
    }, `${kind} result`)
  }

  async #view(kind: 'preview' | 'diff', changeSet: AgentChangeSet<TArtifact>, signal?: AbortSignal): Promise<AgentArtifactView> {
    const envelope = this.#owned(changeSet)
    const identity = await this.#requireCurrent(envelope, signal)
    let raw: unknown
    try { raw = await this.adapter[kind]({ ...this.#context(identity), changeSet: envelope, signal }) }
    catch (error) { throw asAgentToolsError(error) }
    abortIfNeeded(signal)
    const value = adapterJson(raw, `${kind} result`)
    if (!record(value)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result must be an object.`)
    exactAdapterKeys(value, ['data', 'issues', 'evidence'], `${kind} result`)
    if (!Object.prototype.hasOwnProperty.call(value, 'data')) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${kind} result must include data.`)
    return immutableJson({
      protocol: AGENT_TOOLS_PROTOCOL,
      protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
      changeSetId: envelope.changeSetId,
      identity,
      data: value.data,
      issues: normalizeIssues(value.issues),
      evidence: normalizeEvidence(value.evidence),
    }, `${kind} result`)
  }

  async #commit(envelope: Readonly<AgentChangeSetEnvelope>, options: AgentCommitOptions, markMutationStarted: () => void): Promise<AgentCommitResult> {
    abortIfNeeded(options.signal)
    const changeSet = this.#track(envelope)
    const report = await this.validate(changeSet, options.signal)
    if (!report.valid) throw new AgentToolsError('VALIDATION_FAILED', 'Change set validation failed.', false, report.issues)
    const capabilities = await this.capabilities(options.signal)
    const byName = new Map(capabilities.capabilities.map((entry) => [entry.name, entry]))
    const destructive = envelope.operations.filter((operation) => {
      const capability = byName.get(operation.name)
      return capability?.destructive || capability?.requiresConfirmation
    })
    // This is the final optimistic guard. The adapter must repeat it atomically.
    const identity = await this.#requireCurrent(envelope, options.signal)
    if (destructive.length) await this.#confirmMutation('commit', identity, destructive, options.confirmation, envelope)
    abortIfNeeded(options.signal)
    let committed: AgentAdapterCommitResult<TArtifact>
    markMutationStarted()
    try {
      committed = await this.adapter.commit({ ...this.#context(identity), changeSet: envelope, idempotencyKey: options.idempotencyKey, expectedRevision: identity.revision, expectedFingerprint: identity.fingerprint, signal: options.signal })
    } catch (error) { throw asAgentToolsError(error) }
    const checked = await this.#checkReplacement(committed, identity, options.signal)
    // The mutation already succeeded. Adopt only after the adapter's receipt was
    // read back from the exact replacement artifact, before optional proof work.
    this.#artifact = checked.artifact
    this.#identity = checked.identity
    let verified: AgentVerificationResult
    try {
      const verifiedRaw = await this.adapter.verify({ ...this.#contextFor(checked.artifact, checked.identity), changeSet: envelope, stage: 'committed', commit: checked, signal: options.signal })
      verified = verification(envelope.changeSetId, verifiedRaw)
    } catch (error) {
      // The mutation is already authoritative. Never report it as a retryable
      // commit failure merely because optional proof work failed or was aborted.
      const failure = asAgentToolsError(error)
      verified = immutableJson({
        protocol: AGENT_TOOLS_PROTOCOL,
        protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
        changeSetId: envelope.changeSetId,
        verified: false,
        checks: [{ name: 'committed-output', passed: false, message: failure.message }],
        issues: [{ severity: 'error', code: `POST_COMMIT_${failure.code}`, message: failure.message, retryable: failure.retryable }],
      }, 'commit verification failure')
    }
    return immutableJson({
      protocol: AGENT_TOOLS_PROTOCOL,
      protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
      changeSetId: envelope.changeSetId,
      idempotencyKey: options.idempotencyKey,
      identity: checked.identity,
      ...(checked.data !== undefined ? { data: checked.data } : {}),
      ...(checked.evidence ? { evidence: checked.evidence } : {}),
      ...(checked.deduplicated !== undefined ? { deduplicated: checked.deduplicated } : {}),
      verification: verified,
    }, 'commit result')
  }

  async #restore(request: AgentRestoreRequest, markMutationStarted: () => void): Promise<AgentCommitResult> {
    abortIfNeeded(request.signal)
    const identity = await this.#refreshIdentity(request.signal)
    assertExpected(identity, request.expectedRevision, request.expectedFingerprint)
    await this.#confirmMutation('restore', identity, [], request.confirmation)
    let restored: AgentAdapterCommitResult<TArtifact>
    markMutationStarted()
    try {
      restored = await this.adapter.restore!({ ...this.#context(identity), versionId: request.versionId, idempotencyKey: request.idempotencyKey, expectedRevision: identity.revision, expectedFingerprint: request.expectedFingerprint, signal: request.signal })
    } catch (error) { throw asAgentToolsError(error) }
    const checked = await this.#checkReplacement(restored, identity, request.signal)
    this.#artifact = checked.artifact
    this.#identity = checked.identity
    const noChangeSet = `restore-${stableId({ artifactId: identity.artifactId, versionId: request.versionId })}`
    const verified: AgentVerificationResult = immutableJson({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, changeSetId: noChangeSet, verified: true, checks: [{ name: 'adapter-identity', passed: true }], issues: [] }, 'restore verification')
    return immutableJson({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, changeSetId: noChangeSet, idempotencyKey: request.idempotencyKey, identity: checked.identity, ...(checked.data !== undefined ? { data: checked.data } : {}), ...(checked.evidence ? { evidence: checked.evidence } : {}), ...(checked.deduplicated !== undefined ? { deduplicated: checked.deduplicated } : {}), verification: verified }, 'restore result')
  }

  async #checkReplacement(result: AgentAdapterCommitResult<TArtifact>, prior: AgentArtifactIdentity, signal?: AbortSignal): Promise<Readonly<AgentAdapterCommitResult<TArtifact>>> {
    const fields = adapterDataRecord(result, 'Adapter commit result')
    exactAdapterKeys(fields, ['artifact', 'identity', 'data', 'evidence', 'deduplicated'], 'Adapter commit result')
    if (!Object.prototype.hasOwnProperty.call(fields, 'artifact')) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter commit result must include an artifact.')
    if (fields.deduplicated !== undefined && typeof fields.deduplicated !== 'boolean') throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter commit result has an invalid deduplicated field.')
    const checked = adapterJson({ identity: fields.identity, ...(fields.data !== undefined ? { data: fields.data } : {}), ...(fields.evidence !== undefined ? { evidence: normalizeEvidence(fields.evidence) } : {}), ...(fields.deduplicated !== undefined ? { deduplicated: fields.deduplicated } : {}) }, 'adapter commit result')
    if (!record(checked)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter commit result is malformed.')
    const checkedIdentity = adapterIdentity(checked.identity, 'adapter commit identity')
    if (checkedIdentity.artifactId !== prior.artifactId || checkedIdentity.format !== prior.format) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Committed artifact identity changed artifactId or format.')
    let observed: AgentArtifactIdentity
    const artifact = fields.artifact as TArtifact
    try { observed = await this.adapter.identity({ artifact, actor: this.actor, signal }) }
    catch (error) { throw asAgentToolsError(error) }
    const frozenObserved = adapterIdentity(observed, 'committed artifact identity')
    if (!sameIdentity(frozenObserved, checkedIdentity)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter commit identity does not match the replacement artifact.')
    if (frozenObserved.revision === prior.revision) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'A successful mutation must advance the artifact revision.')
    return Object.freeze({ artifact, ...checked, identity: checkedIdentity })
  }

  async #confirmMutation(action: 'commit' | 'restore', identity: AgentArtifactIdentity, destructiveOperations: readonly AgentOperation[], confirmation?: import('./types').JsonValue, changeSet?: Readonly<AgentChangeSetEnvelope>): Promise<void> {
    if (!this.#confirm) throw new AgentToolsError('CONFIRMATION_REQUIRED', `${action} requires a host confirmation hook.`)
    let allowed: unknown
    const safeConfirmation = confirmation === undefined ? undefined : immutableJson(confirmation, 'confirmation') as import('./types').JsonValue
    try { allowed = await this.#confirm({ action, actor: this.actor, identity, changeSet, destructiveOperations: Object.freeze([...destructiveOperations]), confirmation: safeConfirmation }) }
    catch (error) { throw asAgentToolsError(error) }
    if (typeof allowed !== 'boolean') throw new AgentToolsError('CONFIRMATION_DENIED', `${action} confirmation hook did not return a boolean.`)
    if (!allowed) throw new AgentToolsError('CONFIRMATION_DENIED', `${action} was denied by the host.`)
  }

  async #requireCurrent(envelope: Readonly<AgentChangeSetEnvelope>, signal?: AbortSignal): Promise<Readonly<AgentArtifactIdentity>> {
    const identity = await this.#refreshIdentity(signal)
    assertExpected(identity, envelope.baseRevision, envelope.baseFingerprint)
    return identity
  }

  async #refreshIdentity(signal?: AbortSignal): Promise<Readonly<AgentArtifactIdentity>> {
    abortIfNeeded(signal)
    let identity
    try { identity = await this.adapter.identity({ artifact: this.#artifact, actor: this.actor, signal }) }
    catch (error) { throw asAgentToolsError(error) }
    abortIfNeeded(signal)
    const checked = adapterIdentity(identity, 'artifact identity')
    if (checked.artifactId !== this.#identity.artifactId || checked.format !== this.#identity.format) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter changed artifactId or format without a committed replacement.')
    this.#identity = checked
    return checked
  }

  #context(sourceIdentity: Readonly<AgentArtifactIdentity>) { return this.#contextFor(this.#artifact, sourceIdentity) }
  #contextFor(artifact: TArtifact, sourceIdentity: Readonly<AgentArtifactIdentity>) { return { artifact, actor: this.actor, sourceIdentity, limits: this.limits } }

  #owned(changeSet: AgentChangeSet<TArtifact>): Readonly<AgentChangeSetEnvelope> {
    if (!(changeSet instanceof AgentChangeSet) || !this.#changeSets.has(changeSet)) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set does not belong to this session.')
    const { changeSetId: _ignored, protocol: _protocol, protocolVersion: _protocolVersion, ...content } = changeSet.envelope
    if (changeSet.envelope.protocol !== AGENT_CHANGESET_PROTOCOL || changeSet.envelope.protocolVersion !== AGENT_CHANGESET_PROTOCOL_VERSION || changeSet.envelope.changeSetId !== `changeset-${stableId(content)}`) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set envelope is malformed or has been modified.')
    if (changeSet.envelope.operations.length < 1 || changeSet.envelope.operations.length > this.limits.maxOperations || jsonBytes(changeSet.envelope.operations) > this.limits.maxOperationBytes) throw new AgentToolsError('LIMIT_EXCEEDED', 'Change set exceeds this session limits.')
    return changeSet.envelope
  }

  #track(envelope: Readonly<AgentChangeSetEnvelope>): AgentChangeSet<TArtifact> {
    const changeSet = new AgentChangeSet(envelope, this)
    this.#changeSets.add(changeSet)
    return changeSet
  }

  #serialize<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.#mutationTail.then(operation, operation)
    this.#mutationTail = result.then(() => undefined, () => undefined)
    return result
  }
}

export async function createAgentSession<TArtifact>(options: CreateAgentSessionOptions<TArtifact>): Promise<AgentSession<TArtifact>> {
  const actor = immutableJson(options.actor, 'actor')
  validateActor(actor)
  const limits = immutableJson({ ...DEFAULT_AGENT_SESSION_LIMITS, ...options.limits }, 'session limits')
  for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value < 1) throw new AgentToolsError('INVALID_ARGUMENT', `${name} must be a positive safe integer.`)
  const adapter = options.adapter ?? options.registry?.resolve(options.artifact, options.format)
  if (!adapter) throw new AgentToolsError('INVALID_ARGUMENT', 'An adapter or registry is required.')
  if (!adapter.supports(options.artifact)) throw new AgentToolsError('INVALID_ARGUMENT', `Adapter ${adapter.id} does not support the artifact.`)
  let identity
  try { identity = await adapter.identity({ artifact: options.artifact, actor, signal: options.signal }) }
  catch (error) { throw asAgentToolsError(error) }
  abortIfNeeded(options.signal)
  const checked = adapterIdentity(identity, 'artifact identity')
  if (checked.format !== adapter.format) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Adapter format ${adapter.format} does not match artifact format ${checked.format}.`)
  return new AgentSession(options.artifact, checked, actor, adapter, limits, options.confirmDestructive)
}

function verification(changeSetId: string, raw: unknown): AgentVerificationResult {
  const value = adapterJson(raw, 'verification result')
  if (!record(value)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Verification result must be an object.')
  exactAdapterKeys(value, ['verified', 'checks', 'issues'], 'Verification result')
  if (typeof value.verified !== 'boolean' || !Array.isArray(value.checks)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Verification contains an invalid summary or checks list.')
  const checks = value.checks.map((rawCheck, index) => {
    if (!record(rawCheck)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Verification check ${index} must be an object.`)
    exactAdapterKeys(rawCheck, ['name', 'passed', 'message', 'evidence'], `Verification check ${index}`)
    if (typeof rawCheck.name !== 'string' || !rawCheck.name.trim() || typeof rawCheck.passed !== 'boolean' || rawCheck.message !== undefined && typeof rawCheck.message !== 'string') throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Verification check ${index} is malformed.`)
    return {
      name: rawCheck.name,
      passed: rawCheck.passed,
      ...(rawCheck.message !== undefined ? { message: rawCheck.message } : {}),
      ...(rawCheck.evidence !== undefined ? { evidence: normalizeEvidence(rawCheck.evidence) } : {}),
    }
  })
  const issues = normalizeIssues(value.issues)
  if (value.verified !== checks.every(({ passed }) => passed) || value.verified && issues.some(({ severity }) => severity === 'error' || severity === 'refusal')) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Verification summary is inconsistent with its checks or issues.')
  return immutableJson({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, changeSetId, verified: value.verified, checks, issues }, 'verification result')
}

function normalizeIssues(issues: unknown): readonly AgentIssue[] {
  const values = adapterJson(issues, 'adapter issues')
  if (!Array.isArray(values)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter issues must be an array.')
  return values.map((value, index) => {
    if (!record(value)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Issue ${index} must be an object.`)
    exactAdapterKeys(value, ['severity', 'code', 'message', 'path', 'operationId', 'retryable', 'details'], `Issue ${index}`)
    if (!['error', 'warning', 'refusal'].includes(String(value.severity)) || typeof value.code !== 'string' || !value.code.trim() || typeof value.message !== 'string' || !value.message.trim() || value.path !== undefined && typeof value.path !== 'string' || value.operationId !== undefined && typeof value.operationId !== 'string' || value.retryable !== undefined && typeof value.retryable !== 'boolean' || value.details !== undefined && !record(value.details)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Issue ${index} is malformed.`)
    return value as unknown as AgentIssue
  }).sort((a, b) => canonicalJson([a.operationId ?? '', a.path ?? '', a.severity, a.code, a.message]).localeCompare(canonicalJson([b.operationId ?? '', b.path ?? '', b.severity, b.code, b.message])))
}

function normalizeEvidence(evidence: unknown): readonly import('./types').AgentArtifactEvidence[] {
  const values = adapterJson(evidence, 'adapter evidence')
  if (!Array.isArray(values)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', 'Adapter evidence must be an array.')
  return values.map((value, index) => {
    if (!record(value)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Evidence ${index} must be an object.`)
    exactAdapterKeys(value, ['kind', 'description', 'data'], `Evidence ${index}`)
    if (typeof value.kind !== 'string' || !value.kind.trim() || value.description !== undefined && typeof value.description !== 'string') throw new AgentToolsError('INVALID_ADAPTER_RESULT', `Evidence ${index} is malformed.`)
    return value as unknown as import('./types').AgentArtifactEvidence
  })
}

function adapterIdentity(raw: unknown, label: string): Readonly<AgentArtifactIdentity> {
  const value = adapterJson(raw, label)
  if (!record(value)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} must be an object.`)
  exactAdapterKeys(value, ['artifactId', 'format', 'mediaType', 'revision', 'fingerprint'], label)
  if (typeof value.artifactId !== 'string' || !value.artifactId.trim() || typeof value.format !== 'string' || !value.format.trim() || typeof value.revision !== 'string' || !value.revision.trim() || typeof value.fingerprint !== 'string' || !value.fingerprint.trim() || value.mediaType !== undefined && typeof value.mediaType !== 'string') throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} is incomplete or malformed.`)
  return value as unknown as Readonly<AgentArtifactIdentity>
}

function adapterJson(value: unknown, label: string): Readonly<import('./types').JsonValue> {
  try { return immutableJson(value, label) as Readonly<import('./types').JsonValue> }
  catch (error) {
    if (error instanceof AgentToolsError && error.code === 'INVALID_ARGUMENT') throw new AgentToolsError('INVALID_ADAPTER_RESULT', error.message, false, undefined, error.details, { cause: error })
    if (error instanceof AgentToolsError) throw error
    throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} could not be read as plain JSON.`, false, undefined, undefined, { cause: error })
  }
}

function validateActor(actor: unknown): asserts actor is Readonly<AgentActor> {
  if (!record(actor)) throw new AgentToolsError('INVALID_ARGUMENT', 'Agent actor must be a JSON object.')
  exactInputKeys(actor, ['id', 'kind', 'displayName', 'provider', 'model', 'metadata'], 'Agent actor')
  if (typeof actor.id !== 'string' || !validId(actor.id) || actor.kind !== 'agent' || (actor.displayName !== undefined && typeof actor.displayName !== 'string') || (actor.provider !== undefined && typeof actor.provider !== 'string') || (actor.model !== undefined && typeof actor.model !== 'string') || (actor.metadata !== undefined && !record(actor.metadata))) throw new AgentToolsError('INVALID_ARGUMENT', 'Agent actor has malformed fields or lacks a valid stable id and kind `agent`.')
}

function validateEnvelopeShape(envelope: unknown, limits: Readonly<AgentSessionLimits>): asserts envelope is Readonly<AgentChangeSetEnvelope> {
  if (!record(envelope)) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set must be a JSON object.')
  exactInputKeys(envelope, ['protocol', 'protocolVersion', 'changeSetId', 'artifactId', 'format', 'baseRevision', 'baseFingerprint', 'actor', 'operations', 'metadata'], 'Change set')
  if (envelope.protocol !== AGENT_CHANGESET_PROTOCOL || envelope.protocolVersion !== AGENT_CHANGESET_PROTOCOL_VERSION || typeof envelope.changeSetId !== 'string' || !validId(envelope.changeSetId) || typeof envelope.artifactId !== 'string' || !envelope.artifactId.trim() || typeof envelope.format !== 'string' || !envelope.format.trim() || typeof envelope.baseRevision !== 'string' || !envelope.baseRevision.trim() || typeof envelope.baseFingerprint !== 'string' || !envelope.baseFingerprint.trim()) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set envelope has malformed identity fields.')
  validateActor(envelope.actor)
  if (!Array.isArray(envelope.operations) || envelope.operations.length < 1 || envelope.operations.length > limits.maxOperations) throw new AgentToolsError('LIMIT_EXCEEDED', 'Change set operation count exceeds this session limits.')
  const ids = new Set<string>()
  envelope.operations.forEach((raw, index) => {
    if (!record(raw)) throw new AgentToolsError('INVALID_ARGUMENT', `Operation ${index} must be an object.`)
    exactInputKeys(raw, ['operationId', 'name', 'input'], `Operation ${index}`)
    if (typeof raw.operationId !== 'string' || !validId(raw.operationId) || typeof raw.name !== 'string' || !validName(raw.name) || !record(raw.input) || ids.has(raw.operationId)) throw new AgentToolsError('INVALID_ARGUMENT', `Operation ${index} is malformed or duplicated.`)
    ids.add(raw.operationId)
  })
  if (envelope.metadata !== undefined && !record(envelope.metadata)) throw new AgentToolsError('INVALID_ARGUMENT', 'Change set metadata must be a JSON object.')
  if (jsonBytes(envelope.operations) > limits.maxOperationBytes) throw new AgentToolsError('LIMIT_EXCEEDED', 'Change set exceeds this session limits.')
}

function record(value: unknown): value is Record<string, unknown> { return typeof value === 'object' && value !== null && !Array.isArray(value) }
function adapterDataRecord(value: unknown, label: string): Record<string, unknown> {
  if (!record(value)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} must be an object.`)
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== null) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} must use a plain object.`)
  const result: Record<string, unknown> = {}
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key === 'symbol') throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} contains a symbol property.`)
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    if (!descriptor?.enumerable || !('value' in descriptor)) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} properties must be enumerable data properties.`)
    Object.defineProperty(result, key, { configurable: true, enumerable: true, writable: true, value: descriptor.value })
  }
  return result
}
function exactAdapterKeys(value: object, allowed: readonly string[], label: string): void { const extras = Object.keys(value).filter((key) => !allowed.includes(key)); if (extras.length) throw new AgentToolsError('INVALID_ADAPTER_RESULT', `${label} contains unsupported field(s): ${extras.sort().join(', ')}.`) }
function exactInputKeys(value: object, allowed: readonly string[], label: string): void { const extras = Object.keys(value).filter((key) => !allowed.includes(key)); if (extras.length) throw new AgentToolsError('INVALID_ARGUMENT', `${label} contains unsupported field(s): ${extras.sort().join(', ')}.`) }

function baseIdentity(envelope: AgentChangeSetEnvelope): AgentArtifactIdentity { return { artifactId: envelope.artifactId, format: envelope.format, revision: envelope.baseRevision, fingerprint: envelope.baseFingerprint } }
function sameIdentity(left: AgentArtifactIdentity, right: AgentArtifactIdentity): boolean { return left.artifactId === right.artifactId && left.format === right.format && left.revision === right.revision && left.fingerprint === right.fingerprint }
function staleIssue(envelope: AgentChangeSetEnvelope, current: AgentArtifactIdentity): AgentIssue { return { severity: 'error', code: 'STALE_REVISION', message: `Change set is based on revision ${envelope.baseRevision}, but the artifact is at ${current.revision}.`, retryable: true, details: { expectedRevision: envelope.baseRevision, actualRevision: current.revision, expectedFingerprint: envelope.baseFingerprint, actualFingerprint: current.fingerprint } } }
function assertExpected(identity: AgentArtifactIdentity, revision?: string, fingerprint?: string): void { if (revision !== undefined && identity.revision !== revision || fingerprint !== undefined && identity.fingerprint !== fingerprint) throw new AgentToolsError('STALE_REVISION', 'Artifact identity no longer matches the expected revision and fingerprint.', true, undefined, { expectedRevision: revision ?? '', actualRevision: identity.revision, expectedFingerprint: fingerprint ?? '', actualFingerprint: identity.fingerprint }) }
function validName(value: string): boolean { return typeof value === 'string' && /^[a-z][a-z0-9]*(?:[._-][a-z0-9]+)*$/.test(value) && value.length <= 128 }
function validId(value: string): boolean { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:-]{0,191}$/.test(value) }
function validIdempotencyKey(value: string): void { if (!validId(value)) throw new AgentToolsError('INVALID_ARGUMENT', 'idempotencyKey must be 1-192 safe identifier characters.') }
function boundedInteger(requested: number | undefined, maximum: number, name: string): number { const value = requested ?? maximum; if (!Number.isSafeInteger(value) || value < 1 || value > maximum) throw new AgentToolsError('LIMIT_EXCEEDED', `${name} must be between 1 and ${maximum}.`); return value }
function abortIfNeeded(signal?: AbortSignal): void { if (signal?.aborted) throw new AgentToolsError('ABORTED', 'Agent operation was aborted.', true) }
