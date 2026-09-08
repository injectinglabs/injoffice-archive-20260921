import { createDocxAgentAdapter, type DocxAgentArtifact } from '@injoffice/agent-office/docx'
import { AGENT_TOOLS_PROTOCOL, AGENT_TOOLS_PROTOCOL_VERSION, createAgentSession, createAgentToolDispatcher,
  type AgentArtifactView, type AgentBoundedResult, type AgentCapability, type AgentChangeSetEnvelope,
  type AgentCommitResult, type AgentToolCall, type AgentToolCallResult, type AgentToolMethod,
  type AgentValidationReport, type JsonObject } from '@injoffice/agent-tools'
import type { NativeDocxDocumentV1, NativeDocxBlockV1 } from '../../../packages/docs/src/nativeContract'
import type { NativeDocxOfficeMutationEnvelopeV1 } from '../../../packages/docs/src/nativeTransactionAdapterV1'
import type { AgentDemoSession } from './agentDemoRuntime'
import type { AgentDemoArtifact, AgentDemoMode, AgentDemoOperation, AgentDemoScenario } from './agentDemoScenario'
import { DOCX_MEDIA_TYPE, nativeDocxParagraphText } from './docsNativePreview'
import { createBrowserDocxRoundTripRuntime, type DocxRoundTripRuntime } from './docxRoundTripRuntime'
import { createMockDocxProposal, DOCX_MOCK_TEXT_LIMIT, isMockDocxText } from './mockDocxProposal'

export const AGENT_DOCX_NAME = 'northstar-launch-brief.docx'
export const AGENT_DOCX_PROMPT = 'Replace "Northstar Launch Brief" with "Northstar Beta Launch Brief"'
const BOUNDS = { maxItems: 100, maxBytes: 64_000 }
type Target = { targetKind: 'run'; targetId: string; expectedText: string; blockId: string }
function projection(document: NativeDocxDocumentV1, selected?: AgentDemoOperation): Record<string, unknown> {
  const blocks = document.body.blocks.filter((block) => block.paragraph).map((block) => ({ id: block.id, kind: 'paragraph', text: nativeDocxParagraphText(block.paragraph!) }))
  const changed = document.body.blocks.find((block) => block.paragraph?.runs.some((run) => run.id === selected?.input.targetId))
  return { title: blocks[0]?.text ?? 'Northstar Launch Brief', blocks, ...(changed ? { changedBlockId: changed.id } : {}) }
}

/** Native DOCX host: proposals are simulated, package writes and reopened verification are real. */
export async function createNativeDocxAgentSessionInput(mode: AgentDemoMode,
  runtime: DocxRoundTripRuntime = createBrowserDocxRoundTripRuntime(), fetcher: typeof fetch = globalThis.fetch.bind(globalThis)) {
  let disposed = false, nativeWrites = 0, failVerification = false
  let verifiedBytes: Uint8Array | null = null, postCommitReads: number | null = null
  const approvals = new Map<string, string>(), plans = new Map<string, AgentChangeSetEnvelope>()
  const previews = new Map<string, NativeDocxDocumentV1>(), proposalRevisions = new Map<string, string>()
  const traces: Array<{ request: AgentToolCall; response: AgentToolCallResult }> = []
  const proposalKey = (operations: readonly AgentDemoOperation[]) => JSON.stringify(operations.map((item) => [item.op, item.input]))
  const assertOpen = () => { if (disposed) throw new Error('This document session is closed. Reload the sample to continue.') }
  const dispose = () => { if (!disposed) { disposed = true; runtime.terminate(); verifiedBytes = null; approvals.clear(); previews.clear() } }
  try {
    const response = await fetcher(`${import.meta.env.BASE_URL}native-docx/${AGENT_DOCX_NAME}.b64`)
    if (!response.ok) throw new Error(`The sample document could not be loaded (HTTP ${response.status}).`)
    let bytes = Uint8Array.from(atob((await response.text()).trim()), (character) => character.charCodeAt(0))
    let document = (await runtime.extract(bytes.slice())).document
    let selectedOperation: AgentDemoOperation | undefined
    const artifactView = (value = document): AgentDemoArtifact => ({ artifactId: 'northstar-docx-agent-sample', format: 'docx', name: AGENT_DOCX_NAME,
      revision: value.source.package_sha256, content: projection(value, selectedOperation) })
    let mutationQueue = Promise.resolve()
    const serialize = <T>(action: () => Promise<T>): Promise<T> => { const result = mutationQueue.then(action); mutationQueue = result.then(() => undefined, () => undefined); return result }
    const committed = new Map<string, { signature: string; document: NativeDocxDocumentV1 }>()
    const applyCopy = async (envelope: NativeDocxOfficeMutationEnvelopeV1) => {
      assertOpen()
      const source = document
      if (envelope.expected_revision !== source.source.package_sha256) throw new Error('The source revision changed before the native write.')
      const result = await runtime.apply({ original: bytes.slice(), document: structuredClone(source), envelope, sourceName: AGENT_DOCX_NAME })
      const replacement = result.bytes.slice(), reopened = (await runtime.extract(replacement.slice())).document
      assertOpen()
      return { bytes: replacement, document: reopened, sourceRevision: source.source.package_sha256 }
    }
    const host: DocxAgentArtifact = { kind: 'injoffice.docx', artifactId: 'northstar-docx-agent-sample',
      async snapshot() {
        assertOpen()
        // Core checks replacement identity first; the next independent adapter read proves the committed edit.
        if (postCommitReads !== null && postCommitReads++ === 1) { postCommitReads = null; if (failVerification) throw new Error('Demo fault: post-commit verification read unavailable. The write already succeeded.') }
        const reopened = (await runtime.extract(bytes.slice())).document
        assertOpen(); return { document: reopened }
      },
      async preview(envelope) { const candidate = await applyCopy(envelope); previews.set(candidate.document.source.package_sha256, candidate.document); return { document: candidate.document } },
      apply(envelope) { return serialize(async () => {
        assertOpen()
        const signature = JSON.stringify(envelope), retained = committed.get(envelope.mutation_id)
        if (retained) { if (retained.signature !== signature) throw new Error('A different native mutation already used this idempotency key.'); return { snapshot: { document: structuredClone(retained.document) }, deduplicated: true } }
        const candidate = await applyCopy(envelope)
        if (document.source.package_sha256 !== candidate.sourceRevision) throw new Error('The source revision changed during the native write.')
        bytes = candidate.bytes; document = candidate.document; nativeWrites++; verifiedBytes = null; postCommitReads = 0
        committed.set(envelope.mutation_id, { signature, document: structuredClone(document) })
        return { snapshot: { document: structuredClone(document) }, receipt: { nativeWrites }, deduplicated: false }
      }) },
    }
    const core = await createAgentSession({ artifact: host, adapter: createDocxAgentAdapter(), actor: { id: 'playground-local-host', kind: 'agent' },
      confirmDestructive: ({ changeSet, identity }) => !!changeSet && approvals.get(changeSet.changeSetId) === JSON.stringify(changeSet)
        && changeSet.baseRevision === identity.revision && changeSet.baseFingerprint === identity.fingerprint })
    const dispatcher = createAgentToolDispatcher(core)
    let requestNumber = 0
    const call = async <T>(method: AgentToolMethod, params: JsonObject = {}): Promise<T> => {
      assertOpen()
      const request: AgentToolCall = { protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId: `docx-tool-${++requestNumber}`, method, params }
      const result = await dispatcher.dispatch(request)
      assertOpen(); traces.push(structuredClone({ request, response: result }))
      if (!result.ok) throw new Error(`${result.error.code}: ${result.error.message}${result.error.issues?.map((issue) => ` ${issue.code}: ${issue.message}`).join('; ') ?? ''}`)
      return result.result as T
    }
    let disclosed: { revision: string; targets: Target[] } | undefined
    const proposalContext = async (): Promise<JsonObject> => {
      const capabilities = await call<{ capabilities: AgentCapability[] }>('office.capabilities')
      const inspection = await call<AgentBoundedResult>('office.inspect', BOUNDS)
      const blocks: NativeDocxBlockV1[] = []
      let cursor: string | undefined
      do {
        const result = await call<AgentBoundedResult>('office.read', { ...BOUNDS, ...(cursor ? { cursor } : {}) })
        if (result.identity.revision !== inspection.identity.revision) throw new Error('The document changed during proposal discovery.')
        blocks.push(...(result.data as unknown as { blocks: NativeDocxBlockV1[] }).blocks)
        if (blocks.length > 128 || result.truncated && !result.nextCursor) throw new Error('The document exceeds the bounded mock discovery budget.')
        cursor = result.nextCursor
      } while (cursor)
      const targets: Target[] = blocks.flatMap((block) => {
        const paragraph = block.paragraph
        if (!paragraph || paragraph.edit_policy.mode !== 'read-write' || !paragraph.edit_policy.allowed_operations.includes('text.replace')) return []
        return paragraph.runs.filter((run) => run.kind === 'text' && !run.properties?.hidden && isMockDocxText(run.text))
          .map((run) => ({ targetKind: 'run' as const, targetId: run.id, expectedText: run.text!, blockId: block.id }))
      })
      if (targets.length > 128) throw new Error('The document exceeds the bounded mock target budget.')
      disclosed = { revision: inspection.identity.revision, targets }
      return { format: 'docx', identity: { artifactId: inspection.identity.artifactId, revision: inspection.identity.revision, fingerprint: inspection.identity.fingerprint },
        capabilities: capabilities.capabilities.filter((item) => item.name === 'docx.text.replace') as unknown as JsonObject[],
        document: { blocks: blocks.filter((block) => block.paragraph).map((block) => ({ id: block.id, text: nativeDocxParagraphText(block.paragraph!) })) },
        constraints: { maxOperations: 1, maxTextLength: DOCX_MOCK_TEXT_LIMIT, allowedTargets: targets.map((target) => ({ ...target })) } }
    }
    const acceptProposal = (value: unknown): AgentDemoOperation[] => {
      assertOpen()
      const object = (item: unknown): item is Record<string, unknown> => !!item && typeof item === 'object' && !Array.isArray(item)
      if (!disclosed) throw new Error('Disclose the proposal context before accepting a proposal.')
      if (!object(value) || Object.keys(value).some((key) => key !== 'operations') || !Array.isArray(value.operations) || value.operations.length !== 1) throw new Error('A proposal must contain exactly one operation and no other top-level fields.')
      const operation = value.operations[0]
      if (!object(operation) || Object.keys(operation).some((key) => !['name', 'input', 'operationId'].includes(key)) || operation.name !== 'docx.text.replace' || !object(operation.input)) throw new Error('Only docx.text.replace is accepted.')
      const input = operation.input
      if (Object.keys(input).some((key) => !['targetKind', 'targetId', 'expectedText', 'text'].includes(key)) || input.targetKind !== 'run' || !isMockDocxText(input.text) || input.expectedText === input.text) throw new Error('A bounded XML-safe changed text replacement is required.')
      const matches = disclosed.targets.filter((target) => target.targetId === input.targetId && target.expectedText === input.expectedText)
      if (matches.length !== 1) throw new Error('The proposal target or expectedText is outside the disclosed editable runs.')
      if (operation.operationId !== undefined && (typeof operation.operationId !== 'string' || !/^[A-Za-z0-9._:-]{1,128}$/.test(operation.operationId))) throw new Error('The operation id must be a short identifier.')
      const target = matches[0]
      const accepted = [{ operationId: typeof operation.operationId === 'string' ? operation.operationId : 'agent-docx-text', op: 'docx.text.replace', target: target.blockId, value: input.text,
        input: { targetKind: 'run', targetId: target.targetId, expectedText: target.expectedText, text: input.text } }]
      selectedOperation = accepted[0]; proposalRevisions.set(proposalKey(accepted), disclosed.revision)
      return accepted
    }
    const propose = async (request: string) => {
      if (mode === 'refusal') { await proposalContext(); return [{ operationId: 'docx-refusal', op: 'macro.execute', target: 'document', input: {} }] }
      const { capabilities, ...context } = await proposalContext()
      return acceptProposal(createMockDocxProposal({ request, context, capabilities }))
    }
    const session: AgentDemoSession = {
      async capabilities() { const result = await call<{ capabilities: AgentCapability[] }>('office.capabilities'); return { format: 'docx', operations: result.capabilities.map((item) => ({ operation: item.name, access: 'write', destructive: item.destructive })) } },
      async inspect({ selection, maxItems, maxBytes }) { const result = await call<AgentBoundedResult>('office.inspect', { maxItems, maxBytes }); return { artifactId: result.identity.artifactId, revision: result.identity.revision, fingerprint: result.identity.fingerprint, selection, content: projection(document, selectedOperation), truncated: result.truncated } },
      async plan(operations, { expectedRevision }) {
        const proposalRevision = proposalRevisions.get(proposalKey(operations))
        if (proposalRevision && proposalRevision !== expectedRevision) throw new Error('The document changed after proposal discovery. Prepare a new proposal.')
        const envelope = await call<AgentChangeSetEnvelope>('office.plan', { expectedRevision, operations: operations.map((item) => ({ operationId: item.operationId, name: item.op, input: item.input as JsonObject })) })
        plans.set(envelope.changeSetId, envelope)
        const params = { changeSetId: envelope.changeSetId }, before = structuredClone(document), selected = operations[0]
        return { id: envelope.changeSetId, expectedRevision: envelope.baseRevision, operations: structuredClone(operations),
          async validate() { const result = await call<AgentValidationReport>('office.validate', params); return { ok: result.valid, issues: result.issues.map((item) => ({ code: item.code.toLowerCase(), path: item.path ?? '', message: item.message })) } },
          async preview() {
            const result = await call<AgentArtifactView>('office.preview', params)
            const errors = result.issues.filter((item) => item.severity === 'error')
            if (errors.length) throw new Error(`Native preview failed: ${errors.map((item) => `${item.code}: ${item.message}`).join('; ')}. No source write was committed.`)
            const candidate = previews.get(String((result.data as JsonObject).resultingRevision)) ?? before
            return { artifact: { ...artifactView(candidate), content: projection(candidate, selected) }, summary: result.issues.length ? result.issues.map((item) => item.message).join('; ') : 'Native DOCX preview applied to a copy and reopened; source unchanged.', evidence: result.evidence.map((item) => item.description ?? item.kind) }
          },
          async diff() { const result = await call<AgentArtifactView>('office.diff', params); return { changes: result.issues.length ? [] : operations.map((operation) => ({ target: operation.target, before: String(operation.input.expectedText ?? ''), after: String(operation.input.text ?? '') })) } },
          async commit(input) {
            if (input.expectedRevision !== envelope.baseRevision) throw new Error('Commit revision does not match the reviewed change set.')
            const result = await call<AgentCommitResult>('office.commit', { ...params, idempotencyKey: input.idempotencyKey, confirmation: input.confirmation })
            verifiedBytes = result.verification.verified && result.identity.fingerprint === document.source.package_sha256 && result.identity.revision === document.source.package_sha256 ? bytes.slice() : null
            return { artifactId: result.identity.artifactId, previousRevision: envelope.baseRevision, revision: result.identity.revision, fingerprint: result.identity.fingerprint,
              operationIds: envelope.operations.map((item) => item.operationId), evidence: result.evidence?.map((item) => item.description ?? item.kind) ?? [], verification: result.verification, content: projection(document, selected) }
          },
          async verify(receipt) { return { ok: receipt.verification.verified, revision: receipt.revision, fingerprint: receipt.fingerprint, evidence: receipt.verification.checks.map((item) => item.message ?? item.name) } },
        }
      },
    }
    const prompt = mode === 'safe' ? AGENT_DOCX_PROMPT : 'Execute an unsupported document macro.'
    const operations = await propose(prompt)
    const scenario: AgentDemoScenario = { artifact: artifactView(), operations, prompt, summary: mode === 'safe' ? 'Discover an anchored text run, review a native DOCX preview, approve its exact change set, and verify the saved file.' : 'Refuse macro execution before any file write.' }
    return { scenario, operations, session, dispose, propose, proposalContext, acceptProposal,
      async approve(changeSetId: string) {
        assertOpen(); const envelope = plans.get(changeSetId)
        if (!envelope) throw new Error('Only an exact change set retained by this host session can be approved.')
        const current = await call<AgentBoundedResult>('office.inspect', BOUNDS)
        if (current.identity.revision !== envelope.baseRevision || current.identity.fingerprint !== envelope.baseFingerprint) throw new Error('The document changed revision. Prepare and review a new proposal before approval.')
        approvals.set(changeSetId, JSON.stringify(envelope))
      },
      async simulateConcurrentEdit() { await serialize(async () => {
        assertOpen()
        const paragraph = document.body.blocks.find((block) => block.paragraph?.edit_policy.mode === 'read-write' && block.paragraph.edit_policy.allowed_operations.includes('text.replace') && block.paragraph.runs.some((run) => run.kind === 'text'))?.paragraph
        const run = paragraph?.runs.find((item) => item.kind === 'text')
        if (!run) throw new Error('No editable native text run is available for the concurrent edit.')
        const candidate = await applyCopy({ protocol: 'injoffice.office.mutations', version: 1, format: 'docx', mutation_id: 'host-concurrent-docx-edit', expected_revision: document.source.package_sha256,
          payload: { mutations: [{ target_kind: 'run', target_id: run.id, expected_xml_sha256: run.anchor.xml_sha256, text: `${run.text ?? ''} (reviewed)` }] } })
        bytes = candidate.bytes; document = candidate.document; nativeWrites++; verifiedBytes = null; approvals.clear(); postCommitReads = null
      }) },
      setVerificationFailure(enabled: boolean) { assertOpen(); failVerification = enabled },
      trace: () => structuredClone(traces), stats: () => ({ nativeWrites }),
      download: () => !disposed && verifiedBytes ? new Blob([new Uint8Array(verifiedBytes).buffer], { type: DOCX_MEDIA_TYPE }) : null,
    }
  } catch (error) { dispose(); throw error }
}
