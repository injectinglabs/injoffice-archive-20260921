import { AgentToolsError, asAgentToolsError } from './errors'
import { canonicalJson, cloneJson, immutableJson } from './json'
import type { AgentChangeSet, AgentSession } from './session'
import type { AgentOperationInput, AgentToolCallResult, AgentToolDescriptor, AgentToolMethod, JsonObject, JsonValue } from './types'
import { AGENT_TOOLS_PROTOCOL, AGENT_TOOLS_PROTOCOL_VERSION } from './types'

const EMPTY_SCHEMA = schema({}, [])
const CHANGE_SET_ID = { type: 'string', minLength: 1, maxLength: 192 } as JsonObject
const BOUNDS = {
  query: { type: 'object' },
  cursor: { type: 'string' },
  maxItems: { type: 'integer', minimum: 1 },
  maxBytes: { type: 'integer', minimum: 1 },
} satisfies Record<string, JsonObject>

export const AGENT_TOOL_DESCRIPTORS: readonly AgentToolDescriptor[] = immutableJson([
  { name: 'office.capabilities', description: 'List format-specific operations and their JSON input schemas.', inputSchema: EMPTY_SCHEMA },
  { name: 'office.inspect', description: 'Inspect bounded structural context for the current artifact revision.', inputSchema: schema(BOUNDS) },
  { name: 'office.read', description: 'Read bounded artifact content at the current revision.', inputSchema: schema(BOUNDS) },
  { name: 'office.plan', description: 'Create an immutable, revision-bound change set without mutating the artifact.', inputSchema: schema({ operations: { type: 'array', minItems: 1, items: { type: 'object', properties: { operationId: { type: 'string' }, name: { type: 'string' }, input: { type: 'object' } }, required: ['name', 'input'], additionalProperties: false } }, expectedRevision: { type: 'string' }, expectedFingerprint: { type: 'string' }, metadata: { type: 'object' } }, ['operations']) },
  { name: 'office.validate', description: 'Validate a planned change set against core and adapter constraints.', inputSchema: schema({ changeSetId: CHANGE_SET_ID }, ['changeSetId']) },
  { name: 'office.preview', description: 'Produce an adapter-defined preview without mutating the artifact.', inputSchema: schema({ changeSetId: CHANGE_SET_ID }, ['changeSetId']) },
  { name: 'office.diff', description: 'Describe the exact proposed changes without mutating the artifact.', inputSchema: schema({ changeSetId: CHANGE_SET_ID }, ['changeSetId']) },
  { name: 'office.commit', description: 'Atomically commit a valid revision-bound change set and verify the committed artifact.', inputSchema: schema({ changeSetId: CHANGE_SET_ID, idempotencyKey: { type: 'string', minLength: 1, maxLength: 192 }, confirmation: {} }, ['changeSetId', 'idempotencyKey']) },
  { name: 'office.verify', description: 'Verify the planned result. Commit always separately verifies the committed replacement.', inputSchema: schema({ changeSetId: CHANGE_SET_ID }, ['changeSetId']) },
  { name: 'office.restore', description: 'Restore a host version when the selected adapter supports it.', inputSchema: schema({ versionId: { type: 'string', minLength: 1 }, expectedRevision: { type: 'string', minLength: 1 }, expectedFingerprint: { type: 'string' }, idempotencyKey: { type: 'string', minLength: 1, maxLength: 192 }, confirmation: {} }, ['versionId', 'expectedRevision', 'idempotencyKey']) },
] satisfies AgentToolDescriptor[], 'agent tool descriptors')

export interface AgentToolDispatcherOptions { maxChangeSets?: number }

export interface AgentToolDispatcher {
  readonly descriptors: readonly AgentToolDescriptor[]
  readonly changeSetCount: number
  /** Accepts decoded, untrusted JSON and validates the complete tool envelope. */
  dispatch(call: unknown): Promise<AgentToolCallResult>
  forget(changeSetId: string): boolean
}

export function createAgentToolDispatcher<TArtifact>(session: AgentSession<TArtifact>, options: AgentToolDispatcherOptions = {}): AgentToolDispatcher {
  const maxChangeSets = options.maxChangeSets ?? 64
  if (!Number.isSafeInteger(maxChangeSets) || maxChangeSets < 1 || maxChangeSets > 10_000) throw new AgentToolsError('INVALID_ARGUMENT', 'maxChangeSets must be an integer between 1 and 10000.')
  const changeSets = new Map<string, AgentChangeSet<TArtifact>>()
  const descriptors = session.adapter.restore ? AGENT_TOOL_DESCRIPTORS : immutableJson(AGENT_TOOL_DESCRIPTORS.filter(({ name }) => name !== 'office.restore'), 'agent tool descriptors')

  return {
    descriptors,
    get changeSetCount() { return changeSets.size },
    forget(changeSetId) { return changeSets.delete(changeSetId) },
    async dispatch(call) {
      const requestId = safeRequestId(call)
      try {
        const request = cloneJson(call, 'tool call')
        if (!object(request)) throw new AgentToolsError('INVALID_ARGUMENT', 'Tool call must be a JSON object.')
        if (!object(request.params)) throw new AgentToolsError('INVALID_ARGUMENT', 'Tool call params must be a JSON object.')
        exactKeys(request, ['protocol', 'protocolVersion', 'requestId', 'method', 'params'])
        if (request.protocol !== AGENT_TOOLS_PROTOCOL || request.protocolVersion !== AGENT_TOOLS_PROTOCOL_VERSION) throw new AgentToolsError('INVALID_ARGUMENT', 'Unsupported agent tools protocol or version.')
        if (!requestId.trim() || requestId.length > 192) throw new AgentToolsError('INVALID_ARGUMENT', 'requestId must be 1-192 characters.')
        if (!descriptors.some(({ name }) => name === request.method)) throw new AgentToolsError('INVALID_ARGUMENT', `Unknown or unavailable tool method ${String(request.method)}.`)
        const result = await execute(request.method as AgentToolMethod, request.params)
        return immutableJson({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId, ok: true as const, result }, 'tool result')
      } catch (error) {
        return immutableJson({ protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId, ok: false as const, error: asAgentToolsError(error).toJSON() }, 'tool error')
      }
    },
  }

  async function execute(method: AgentToolMethod, params: JsonObject): Promise<JsonValue> {
    switch (method) {
      case 'office.capabilities': exactKeys(params, []); return asJson(await session.capabilities())
      case 'office.inspect': return asJson(await session.inspect(boundedParams(params)))
      case 'office.read': return asJson(await session.read(boundedParams(params)))
      case 'office.plan': {
        exactKeys(params, ['operations', 'expectedRevision', 'expectedFingerprint', 'metadata'])
        if (!Array.isArray(params.operations)) throw new AgentToolsError('INVALID_ARGUMENT', 'office.plan operations must be an array.')
        optionalString(params, 'expectedRevision', 'office.plan')
        optionalString(params, 'expectedFingerprint', 'office.plan')
        if (params.metadata !== undefined && !object(params.metadata)) throw new AgentToolsError('INVALID_ARGUMENT', 'office.plan metadata must be an object.')
        const changeSet = await session.plan(params.operations as unknown as AgentOperationInput[], {
          ...(typeof params.expectedRevision === 'string' ? { expectedRevision: params.expectedRevision } : {}),
          ...(typeof params.expectedFingerprint === 'string' ? { expectedFingerprint: params.expectedFingerprint } : {}),
          ...(object(params.metadata) ? { metadata: params.metadata } : {}),
        })
        const retained = changeSets.get(changeSet.envelope.changeSetId)
        if (retained && canonicalJson(retained.envelope) !== canonicalJson(changeSet.envelope)) throw new AgentToolsError('INVALID_ARGUMENT', 'A different retained change set has the same compact id.')
        if (!retained && changeSets.size >= maxChangeSets) throw new AgentToolsError('LIMIT_EXCEEDED', `Dispatcher retains at most ${maxChangeSets} change sets; forget one before planning another.`)
        changeSets.set(changeSet.envelope.changeSetId, changeSet)
        return asJson(changeSet.envelope)
      }
      case 'office.validate': return asJson(await selected(params).validate())
      case 'office.preview': return asJson(await selected(params).preview())
      case 'office.diff': return asJson(await selected(params).diff())
      case 'office.verify': return asJson(await selected(params).verify())
      case 'office.commit': {
        exactKeys(params, ['changeSetId', 'idempotencyKey', 'confirmation'])
        if (typeof params.idempotencyKey !== 'string') throw new AgentToolsError('INVALID_ARGUMENT', 'office.commit idempotencyKey is required.')
        return asJson(await selected(params, false).commit({ idempotencyKey: params.idempotencyKey, ...(params.confirmation !== undefined ? { confirmation: params.confirmation } : {}) }))
      }
      case 'office.restore': {
        exactKeys(params, ['versionId', 'expectedRevision', 'expectedFingerprint', 'idempotencyKey', 'confirmation'])
        if (typeof params.versionId !== 'string' || typeof params.expectedRevision !== 'string' || typeof params.idempotencyKey !== 'string') throw new AgentToolsError('INVALID_ARGUMENT', 'office.restore requires versionId, expectedRevision, and idempotencyKey.')
        optionalString(params, 'expectedFingerprint', 'office.restore')
        return asJson(await session.restore({ versionId: params.versionId, expectedRevision: params.expectedRevision, idempotencyKey: params.idempotencyKey, ...(typeof params.expectedFingerprint === 'string' ? { expectedFingerprint: params.expectedFingerprint } : {}), ...(params.confirmation !== undefined ? { confirmation: params.confirmation } : {}) }))
      }
    }
  }

  function selected(params: JsonObject, strict = true): AgentChangeSet<TArtifact> {
    if (strict) exactKeys(params, ['changeSetId'])
    if (typeof params.changeSetId !== 'string') throw new AgentToolsError('INVALID_ARGUMENT', 'changeSetId is required.')
    const found = changeSets.get(params.changeSetId)
    if (!found) throw new AgentToolsError('INVALID_ARGUMENT', `Change set ${params.changeSetId} is not retained by this dispatcher.`)
    return found
  }
}

function schema(properties: Record<string, JsonObject>, required: readonly string[] = []): import('./types').AgentInputSchema {
  return { type: 'object', properties, ...(required.length ? { required: [...required] } : {}), additionalProperties: false }
}

function boundedParams(params: JsonObject): { query?: JsonObject; cursor?: string; maxItems?: number; maxBytes?: number } {
  exactKeys(params, ['query', 'cursor', 'maxItems', 'maxBytes'])
  if ((params.query !== undefined && !object(params.query)) || (params.cursor !== undefined && typeof params.cursor !== 'string') || (params.maxItems !== undefined && typeof params.maxItems !== 'number') || (params.maxBytes !== undefined && typeof params.maxBytes !== 'number')) throw new AgentToolsError('INVALID_ARGUMENT', 'Bounded request has invalid parameter types.')
  return { ...(object(params.query) ? { query: params.query } : {}), ...(typeof params.cursor === 'string' ? { cursor: params.cursor } : {}), ...(typeof params.maxItems === 'number' ? { maxItems: params.maxItems } : {}), ...(typeof params.maxBytes === 'number' ? { maxBytes: params.maxBytes } : {}) }
}

function exactKeys(value: JsonObject, allowed: readonly string[]): void {
  const extras = Object.keys(value).filter((key) => !allowed.includes(key))
  if (extras.length) throw new AgentToolsError('INVALID_ARGUMENT', `Unsupported parameter(s): ${extras.sort().join(', ')}.`)
}
function object(value: unknown): value is JsonObject { return typeof value === 'object' && value !== null && !Array.isArray(value) }
function asJson(value: unknown): JsonValue { return cloneJson(value, 'tool result') as JsonValue }
function optionalString(params: JsonObject, key: string, method: string): void { if (params[key] !== undefined && typeof params[key] !== 'string') throw new AgentToolsError('INVALID_ARGUMENT', `${method} ${key} must be a string.`) }
function safeRequestId(call: unknown): string {
  if (typeof call !== 'object' || call === null || Array.isArray(call)) return ''
  const descriptor = Object.getOwnPropertyDescriptor(call, 'requestId')
  return descriptor && 'value' in descriptor && typeof descriptor.value === 'string' ? descriptor.value : ''
}
