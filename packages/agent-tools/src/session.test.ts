import { describe, expect, it, vi } from 'vitest'
import {
  AGENT_CHANGESET_PROTOCOL,
  AGENT_TOOLS_PROTOCOL,
  AGENT_TOOLS_PROTOCOL_VERSION,
  AgentAdapterRegistry,
  AgentChangeSet,
  AgentToolsError,
  createAgentSession,
  createAgentToolDispatcher,
} from './index'
import type { AgentArtifactAdapter, AgentCapability, AgentToolCall, JsonObject } from './index'

interface Artifact { id: string; revision: number; value: number }

const capabilities: AgentCapability[] = [
  { name: 'value.delete', description: 'Delete the value.', destructive: true, requiresConfirmation: true, inputSchema: { type: 'object', properties: {}, additionalProperties: false } },
  { name: 'value.set', description: 'Set the value.', inputSchema: { type: 'object', properties: { value: { type: 'number' } }, required: ['value'], additionalProperties: false } },
]

function identity(artifact: Artifact) {
  return { artifactId: artifact.id, format: 'test', mediaType: 'application/x-test', revision: String(artifact.revision), fingerprint: `value:${artifact.value}` }
}

function adapter(overrides: Partial<AgentArtifactAdapter<Artifact>> = {}): AgentArtifactAdapter<Artifact> {
  return {
    id: 'test-adapter',
    format: 'test',
    supports: (value): value is Artifact => typeof value === 'object' && value !== null && 'revision' in value,
    identity: async ({ artifact }) => identity(artifact),
    capabilities: async () => [...capabilities].reverse(),
    inspect: async ({ artifact, maxItems }) => ({ data: { revision: artifact.revision }, itemCount: Math.min(1, maxItems), truncated: false }),
    read: async ({ artifact, maxItems }) => ({ data: [{ value: artifact.value }].slice(0, maxItems), itemCount: Math.min(1, maxItems), truncated: false }),
    validate: async ({ changeSet }) => changeSet.operations.flatMap((operation) => operation.name === 'value.set' && typeof operation.input.value !== 'number' ? [{ severity: 'error' as const, code: 'INVALID_VALUE', message: 'value must be numeric', operationId: operation.operationId }] : []),
    preview: async ({ artifact, changeSet }) => ({ data: { value: apply(artifact, changeSet.operations) }, issues: [], evidence: [{ kind: 'preview' }] }),
    diff: async ({ artifact, changeSet }) => ({ data: { before: artifact.value, after: apply(artifact, changeSet.operations) }, issues: [], evidence: [] }),
    verify: async ({ artifact, changeSet, stage }) => {
      const passed = stage === 'planned' || artifact.value === apply({ ...artifact, value: Number(changeSet.operations[0]?.input.previous ?? 0) }, changeSet.operations)
      return { verified: passed, checks: [{ name: `${stage}-value`, passed }], issues: [] }
    },
    commit: async ({ artifact, changeSet, expectedRevision, expectedFingerprint }) => {
      if (identity(artifact).revision !== expectedRevision || identity(artifact).fingerprint !== expectedFingerprint) throw new AgentToolsError('STALE_REVISION', 'adapter CAS failed')
      const replacement = { ...artifact, revision: artifact.revision + 1, value: apply(artifact, changeSet.operations) }
      return { artifact: replacement, identity: identity(replacement), data: { saved: true } }
    },
    restore: async ({ artifact }) => {
      const replacement = { ...artifact, revision: artifact.revision + 1, value: 0 }
      return { artifact: replacement, identity: identity(replacement) }
    },
    ...overrides,
  }
}

function apply(artifact: Artifact, operations: readonly { name: string; input: JsonObject }[]): number {
  let value = artifact.value
  for (const operation of operations) {
    if (operation.name === 'value.set') value = operation.input.value as number
    if (operation.name === 'value.delete') value = 0
  }
  return value
}

async function session(options: Parameters<typeof createAgentSession<Artifact>>[0] extends infer T ? Partial<T> : never = {}) {
  return createAgentSession({
    artifact: { id: 'artifact-1', revision: 1, value: 7 },
    adapter: adapter(),
    actor: { id: 'agent-1', kind: 'agent', displayName: 'Test agent' },
    ...options,
  })
}

describe('agent session', () => {
  it('advertises sorted JSON-schema capabilities and returns bounded reads', async () => {
    const subject = await session({ limits: { maxReadItems: 2, maxReadBytes: 1_024 } })
    const advertised = await subject.capabilities()
    expect(advertised.capabilities.map(({ name }) => name)).toEqual(['value.delete', 'value.set'])
    expect(advertised.capabilities.every(({ inputSchema }) => inputSchema.type === 'object')).toBe(true)
    expect(await subject.read({ maxItems: 1, maxBytes: 100 })).toMatchObject({ appliedLimits: { maxItems: 1, maxBytes: 100 }, itemCount: 1 })
    await expect(subject.read({ maxItems: 3 })).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' })
  })

  it('creates deterministic immutable plans without changing caller input', async () => {
    const subject = await session()
    const input = { name: 'value.set', input: { value: 9 } }
    const first = await subject.plan([input])
    const second = await subject.plan([{ input: { value: 9 }, name: 'value.set' }])
    expect(first.envelope.changeSetId).toBe(second.envelope.changeSetId)
    expect(first.envelope).toMatchObject({ protocol: AGENT_CHANGESET_PROTOCOL, baseRevision: '1', operations: [{ operationId: 'op-0001' }] })
    expect(Object.isFrozen(first.envelope.operations[0]?.input)).toBe(true)
    expect(input).toEqual({ name: 'value.set', input: { value: 9 } })
    expect(() => { (first.envelope.operations[0]!.input as { value: number }).value = 12 }).toThrow()
  })

  it('combines unsupported-operation refusals with adapter validation', async () => {
    const subject = await session()
    const changeSet = await subject.plan([{ name: 'other.operation', input: {} }, { name: 'value.set', input: { value: 'nine' } }])
    const result = await changeSet.validate()
    expect(result.valid).toBe(false)
    expect(result.issues.map(({ code }) => code)).toEqual(['UNSUPPORTED_OPERATION', 'INVALID_VALUE'])
    await expect(changeSet.commit({ idempotencyKey: 'bad-1' })).rejects.toMatchObject({ code: 'VALIDATION_FAILED' })
  })

  it('rejects stale plans before preview or commit', async () => {
    const artifact = { id: 'artifact-1', revision: 1, value: 7 }
    const subject = await createAgentSession({ artifact, adapter: adapter(), actor: { id: 'agent-1', kind: 'agent' } })
    const changeSet = await subject.plan([{ name: 'value.set', input: { value: 8 } }])
    artifact.revision = 2
    await expect(changeSet.preview()).rejects.toMatchObject({ code: 'STALE_REVISION' })
    await expect(changeSet.commit({ idempotencyKey: 'stale-1' })).rejects.toMatchObject({ code: 'VALIDATION_FAILED' })
  })

  it('commits once per key, adopts immutable replacements, and verifies committed output', async () => {
    const commit = vi.fn(adapter().commit)
    const verify = vi.fn(adapter().verify)
    const subject = await session({ adapter: adapter({ commit, verify }) })
    const changeSet = await subject.plan([{ name: 'value.set', input: { value: 11, previous: 7 } }])
    const [first, duplicate] = await Promise.all([
      changeSet.commit({ idempotencyKey: 'same-key' }),
      changeSet.commit({ idempotencyKey: 'same-key' }),
    ])
    expect(first).toEqual(duplicate)
    expect(commit).toHaveBeenCalledTimes(1)
    expect(subject.artifact).toEqual({ id: 'artifact-1', revision: 2, value: 11 })
    expect(first.verification).toMatchObject({ verified: true, checks: [{ name: 'committed-value', passed: true }] })
    expect(verify.mock.calls.at(-1)?.[0]).toMatchObject({ stage: 'committed', artifact: { revision: 2, value: 11 }, sourceIdentity: { revision: '2' } })
    const another = await subject.plan([{ name: 'value.set', input: { value: 12 } }])
    expect(() => another.commit({ idempotencyKey: 'same-key' })).toThrowError(AgentToolsError)
  })

  it('requires and honors host confirmation for destructive commit and restore', async () => {
    const destructive = await session()
    const removal = await destructive.plan([{ name: 'value.delete', input: {} }])
    await expect(removal.commit({ idempotencyKey: 'delete-1' })).rejects.toMatchObject({ code: 'CONFIRMATION_REQUIRED' })

    const confirm = vi.fn(async ({ confirmation }) => confirmation === 'approved')
    const approved = await session({ confirmDestructive: confirm })
    const approvedRemoval = await approved.plan([{ name: 'value.delete', input: {} }])
    await expect(approvedRemoval.commit({ idempotencyKey: 'delete-2', confirmation: 'no' })).rejects.toMatchObject({ code: 'CONFIRMATION_DENIED' })
    await expect(approvedRemoval.commit({ idempotencyKey: 'delete-2', confirmation: 'approved' })).resolves.toMatchObject({ identity: { revision: '2' } })
    await expect(approved.restore({ versionId: 'v1', expectedRevision: '2', idempotencyKey: 'restore-1', confirmation: 'approved' })).resolves.toMatchObject({ identity: { revision: '3' } })
    expect(confirm).toHaveBeenCalledTimes(3)
  })

  it('fails closed when a JavaScript confirmation hook returns a truthy non-boolean', async () => {
    const subject = await session({ confirmDestructive: (async () => 'approved') as never })
    const removal = await subject.plan([{ name: 'value.delete', input: {} }])
    await expect(removal.commit({ idempotencyKey: 'invalid-confirmation-hook', confirmation: 'approved' })).rejects.toMatchObject({ code: 'CONFIRMATION_DENIED' })
  })

  it('refuses oversized or malformed adapter output', async () => {
    const subject = await session({ adapter: adapter({ read: async () => ({ data: 'x'.repeat(100), itemCount: 1, truncated: false }) }) })
    await expect(subject.read({ maxBytes: 10 })).rejects.toMatchObject({ code: 'INVALID_ADAPTER_RESULT' })
    await expect(subject.plan([{ name: 'value.set', input: { value: Number.NaN } }])).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' })
  })

  it('rejects adapter envelope shadowing and malformed nested result fields', async () => {
    const read = async () => ({
      data: { safe: true },
      itemCount: 1,
      truncated: false,
      protocol: 'attacker-controlled',
    }) as unknown as ReturnType<AgentArtifactAdapter<Artifact>['read']> extends Promise<infer T> ? T : never
    const subject = await session({ adapter: adapter({ read }) })
    await expect(subject.read()).rejects.toMatchObject({ code: 'INVALID_ADAPTER_RESULT' })

    const preview = async () => ({ data: {}, issues: [], evidence: [], changeSetId: 'forged' }) as never
    const previewSubject = await session({ adapter: adapter({ preview }) })
    const changeSet = await previewSubject.plan([{ name: 'value.set', input: { value: 8 } }])
    await expect(changeSet.preview()).rejects.toMatchObject({ code: 'INVALID_ADAPTER_RESULT' })

    const malformedIssues = await session({ adapter: adapter({ validate: async () => [null] as never }) })
    const malformedPlan = await malformedIssues.plan([{ name: 'value.set', input: { value: 8 } }])
    await expect(malformedPlan.validate()).rejects.toMatchObject({ code: 'INVALID_ADAPTER_RESULT' })

    let getterCalls = 0
    const maliciousCommit = async () => {
      const result: Record<string, unknown> = { artifact: { id: 'artifact-1', revision: 2, value: 8 } }
      Object.defineProperty(result, 'identity', { enumerable: true, get: () => { getterCalls += 1; return identity(result.artifact as Artifact) } })
      return result as never
    }
    const commitSubject = await session({ adapter: adapter({ commit: maliciousCommit }) })
    const commitPlan = await commitSubject.plan([{ name: 'value.set', input: { value: 8 } }])
    await expect(commitPlan.commit({ idempotencyKey: 'malicious-result' })).rejects.toMatchObject({ code: 'INVALID_ADAPTER_RESULT' })
    expect(getterCalls).toBe(0)
  })

  it('reports post-commit verification failures without making a successful mutation look retryable', async () => {
    const commit = vi.fn(adapter().commit)
    const subject = await session({ adapter: adapter({ commit, verify: async () => { throw Object.assign(new Error('proof timed out'), { name: 'AbortError' }) } }) })
    const changeSet = await subject.plan([{ name: 'value.set', input: { value: 11 } }])
    const first = await changeSet.commit({ idempotencyKey: 'proof-failure' })
    const duplicate = await changeSet.commit({ idempotencyKey: 'proof-failure' })
    expect(commit).toHaveBeenCalledTimes(1)
    expect(subject.identity.revision).toBe('2')
    expect(first).toEqual(duplicate)
    expect(first.verification).toMatchObject({ verified: false, checks: [{ passed: false }], issues: [{ code: 'POST_COMMIT_ABORTED' }] })
  })

  it('does not replay an idempotency key after an ambiguous adapter commit failure', async () => {
    const commit = vi.fn(async () => { throw new Error('connection dropped after write') })
    const subject = await session({ adapter: adapter({ commit }) })
    const changeSet = await subject.plan([{ name: 'value.set', input: { value: 11 } }])
    await expect(changeSet.commit({ idempotencyKey: 'ambiguous-write' })).rejects.toMatchObject({ code: 'ADAPTER_ERROR' })
    await expect(changeSet.commit({ idempotencyKey: 'ambiguous-write' })).rejects.toMatchObject({ code: 'ADAPTER_ERROR' })
    expect(commit).toHaveBeenCalledTimes(1)
  })

  it('observes aborts after an adapter that ignores its signal returns', async () => {
    let release!: (value: { data: JsonObject; itemCount: number; truncated: boolean }) => void
    const subject = await session({ adapter: adapter({ read: async () => new Promise((resolve) => { release = resolve }) }) })
    const controller = new AbortController()
    const pending = subject.read({ signal: controller.signal })
    await vi.waitFor(() => expect(release).toBeTypeOf('function'))
    controller.abort()
    release({ data: {}, itemCount: 0, truncated: false })
    await expect(pending).rejects.toMatchObject({ code: 'ABORTED' })
  })

  it('rejects foreign change-set instances and malformed loaded envelopes', async () => {
    const first = await session()
    const second = await session()
    const planned = await first.plan([{ name: 'value.set', input: { value: 8 } }])
    await expect(second.validate(planned)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' })
    const forged = new AgentChangeSet(planned.envelope, second)
    await expect(second.validate(forged)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' })

    const malformed = structuredClone(planned.envelope) as unknown as { operations: unknown[] }
    malformed.operations = [null]
    expect(() => second.loadChangeSet(malformed as never)).toThrowError(AgentToolsError)
  })
})

describe('registry and JSON dispatcher', () => {
  it('resolves explicit format adapters and refuses ambiguity', () => {
    const registry = new AgentAdapterRegistry([adapter(), adapter({ id: 'second', format: 'other' })])
    expect(registry.resolve({ id: 'a', revision: 1, value: 1 }, 'test').id).toBe('test-adapter')
    expect(() => registry.resolve({ id: 'a', revision: 1, value: 1 })).toThrow(/Multiple/)
  })

  it('dispatches an end-to-end JSON workflow and bounds retained plans', async () => {
    const subject = await session()
    const dispatcher = createAgentToolDispatcher(subject, { maxChangeSets: 1 })
    expect(dispatcher.descriptors.every(({ inputSchema }) => inputSchema.type === 'object' && inputSchema.additionalProperties === false)).toBe(true)
    const invoke = (method: AgentToolCall['method'], params: JsonObject, requestId = method) => dispatcher.dispatch({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId, method, params })
    const planned = await invoke('office.plan', { operations: [{ name: 'value.set', input: { value: 10, previous: 7 } }] })
    expect(planned.ok).toBe(true)
    const changeSetId = planned.ok ? (planned.result as unknown as { changeSetId: string }).changeSetId : ''
    expect(await invoke('office.preview', { changeSetId })).toMatchObject({ ok: true, result: { data: { value: 10 } } })
    expect(await invoke('office.commit', { changeSetId, idempotencyKey: 'wire-commit' })).toMatchObject({ ok: true, result: { verification: { verified: true } } })
    expect(await invoke('office.commit', { changeSetId, idempotencyKey: 'wire-commit' })).toMatchObject({ ok: true })
    const overflow = await invoke('office.plan', { operations: [{ name: 'value.set', input: { value: 12 } }] })
    expect(overflow).toMatchObject({ ok: false, error: { code: 'LIMIT_EXCEEDED' } })
    expect(dispatcher.forget(changeSetId)).toBe(true)
    expect(await invoke('office.plan', { operations: [{ name: 'value.set', input: { value: 12 } }] })).toMatchObject({ ok: true })
  })

  it('returns structured errors for unknown fields and malformed JSON values', async () => {
    const dispatcher = createAgentToolDispatcher(await session())
    const extra = await dispatcher.dispatch({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId: 'bad', method: 'office.inspect', params: { surprise: true } })
    expect(extra).toMatchObject({ ok: false, error: { code: 'INVALID_ARGUMENT', retryable: false } })
    const cyclic: JsonObject = {}
    cyclic.self = cyclic
    const malformed = await dispatcher.dispatch({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId: 'cycle', method: 'office.inspect', params: cyclic })
    expect(malformed).toMatchObject({ ok: false, error: { code: 'INVALID_ARGUMENT' } })
  })

  it('rejects non-JSON JavaScript graphs without invoking accessors', async () => {
    const dispatcher = createAgentToolDispatcher(await session())
    const base = { protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId: 'unsafe', method: 'office.inspect' }
    let getterCalls = 0
    const accessorCall: Record<string, unknown> = { ...base }
    Object.defineProperty(accessorCall, 'params', { enumerable: true, get: () => { getterCalls += 1; return {} } })
    await expect(dispatcher.dispatch(accessorCall)).resolves.toMatchObject({ ok: false, error: { code: 'INVALID_ARGUMENT' } })
    expect(getterCalls).toBe(0)

    const symbolCall = { ...base, params: {} }
    Object.defineProperty(symbolCall.params, Symbol('hidden'), { enumerable: true, value: true })
    await expect(dispatcher.dispatch(symbolCall)).resolves.toMatchObject({ ok: false, error: { code: 'INVALID_ARGUMENT' } })

    const sparseOperations = new Array(1)
    await expect(dispatcher.dispatch({ ...base, requestId: 'sparse', method: 'office.plan', params: { operations: sparseOperations } })).resolves.toMatchObject({ ok: false, error: { code: 'INVALID_ARGUMENT' } })
  })

  it('normalizes negative zero and safely preserves prototype-shaped JSON keys', async () => {
    const inspect = async ({ query }: { query?: JsonObject }) => ({ data: query ?? {}, itemCount: 1, truncated: false })
    const subject = await session({ adapter: adapter({ inspect: inspect as AgentArtifactAdapter<Artifact>['inspect'] }) })
    const dispatcher = createAgentToolDispatcher(subject)
    const invoke = (method: AgentToolCall['method'], params: JsonObject, requestId = method) => dispatcher.dispatch({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId, method, params })
    const planned = await invoke('office.plan', { operations: [{ name: 'value.set', input: { value: -0 } }] })
    expect(planned).toMatchObject({ ok: true, result: { operations: [{ input: { value: 0 } }] } })
    if (planned.ok) {
      const value = (planned.result as unknown as { operations: Array<{ input: { value: number } }> }).operations[0]!.input.value
      expect(Object.is(value, -0)).toBe(false)
    }

    const query = JSON.parse('{"__proto__":{"polluted":true},"constructor":"kept"}') as JsonObject
    const inspected = await invoke('office.inspect', { query })
    expect(inspected.ok).toBe(true)
    if (inspected.ok) {
      const data = (inspected.result as unknown as { data: Record<string, unknown> }).data
      expect(Object.prototype.hasOwnProperty.call(data, '__proto__')).toBe(true)
      expect(data.constructor).toBe('kept')
      expect((Object.prototype as { polluted?: boolean }).polluted).toBeUndefined()
    }
  })

  it('rejects wrong optional wire types instead of silently dropping them', async () => {
    const dispatcher = createAgentToolDispatcher(await session())
    const call = (params: JsonObject) => dispatcher.dispatch({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId: 'wrong-type', method: 'office.plan', params })
    await expect(call({ operations: [{ name: 'value.set', input: { value: 1 } }], expectedRevision: 1 })).resolves.toMatchObject({ ok: false, error: { code: 'INVALID_ARGUMENT' } })
    await expect(call({ operations: [{ name: 'value.set', input: { value: 1 } }], metadata: [] })).resolves.toMatchObject({ ok: false, error: { code: 'INVALID_ARGUMENT' } })
  })

  it('keeps per-dispatcher descriptors immutable when restore is unavailable', async () => {
    const withoutRestore = adapter()
    delete withoutRestore.restore
    const dispatcher = createAgentToolDispatcher(await session({ adapter: withoutRestore }))
    expect(Object.isFrozen(dispatcher.descriptors)).toBe(true)
    expect(() => (dispatcher.descriptors as AgentCapability[]).pop()).toThrow()
  })
})
