import { createNativePptxAgentAdapter, type NativePptxAgentArtifact, type PptxNativeMutationRequest, type PptxNativeExactParagraph } from '@injoffice/agent-office/pptx'
import { AGENT_TOOLS_PROTOCOL, AGENT_TOOLS_PROTOCOL_VERSION, createAgentSession, createAgentToolDispatcher,
  type AgentArtifactView, type AgentBoundedResult, type AgentCapability, type AgentChangeSetEnvelope, type AgentCommitResult,
  type AgentToolCall, type AgentToolCallResult, type AgentToolMethod, type AgentValidationReport, type JsonObject } from '@injoffice/agent-tools'
import { assertNativePptx, type NativeElement, type NativePptxDeck } from '@injoffice/pptx-native'
import type { AgentDemoSession } from './agentDemoRuntime'
import type { AgentDemoArtifact, AgentDemoMode, AgentDemoOperation, AgentDemoScenario } from './agentDemoScenario'
import { createBrowserPptxRoundTripRuntime, type PptxRoundTripRuntime } from './pptxRoundTripRuntime'
import { createMockPptxProposal, type MockPptxTarget } from './mockPptxProposal'

export const AGENT_PPTX_NAME = 'northstar-quarterly-review.pptx'
const SAMPLE_PATH = new URL('../../../go/pptxpatch/testdata/playground_northstar_review.pptx', import.meta.url).href
const MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const BOUNDS = { maxItems: 100, maxBytes: 64_000 }
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value)
function elements(deck: NativePptxDeck) { const result: Array<{ slideId: string; element: NativeElement }> = []; const visit = (slideId: string, element: NativeElement) => { result.push({ slideId, element }); if (element.kind === 'group') element.children.forEach((child) => visit(slideId, child)) }; deck.slides.forEach((slide) => slide.elements.forEach((element) => visit(slide.id, element))); return result }
function firstText(element?: NativeElement) { return element?.kind === 'text' || element?.kind === 'shape' ? element.paragraphs[0]?.runs[0]?.text ?? '' : '' }
function exactTarget(element: NativeElement): boolean {
  return (element.kind === 'text' || element.kind === 'shape') && !!element.source && element.compatibility.status !== 'refused' &&
    !!firstText(element) && firstText(element).length <= 1000 && !/["\t\r\n]/.test(firstText(element)) && element.paragraphs.every((paragraph) => paragraph.bullet !== true && paragraph.runs.every((run) =>
      run.bold !== undefined && run.italic !== undefined && run.fontSizeHundredthPt !== undefined && !!run.color && /^[0-9A-F]{6}$/.test(run.color) && !!run.fontFamily))
}
function projection(deck: NativePptxDeck, operation?: AgentDemoOperation): Record<string, unknown> {
  const all = elements(deck)
  const slideId = all.find(({ element }) => element.id === operation?.input.elementId)?.slideId ?? deck.slides[0]?.id
  const blocks = all.filter((item) => item.slideId === slideId && (item.element.kind === 'text' || item.element.kind === 'shape'))
    .map(({ element }) => ({ id: element.id, text: element.kind === 'text' || element.kind === 'shape' ? element.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n') : '' })).filter((block) => block.text)
  return { title: blocks[0]?.text ?? 'Presentation', subtitle: `Slide ${deck.slides.findIndex((slide) => slide.id === slideId) + 1} of ${deck.slides.length}`, blocks,
    ...(operation ? { changedBlockId: operation.input.elementId } : {}) }
}

/** Tab-owned native bytes and trusted approvals; the proposer receives only bounded text context. */
export async function createNativePptxAgentSessionInput(mode: AgentDemoMode,
  runtime: PptxRoundTripRuntime = createBrowserPptxRoundTripRuntime(), fetcher: typeof fetch = globalThis.fetch.bind(globalThis)) {
  let disposed = false, nativeWrites = 0, failVerification = false
  let postCommitReads: number | null = null
  let verifiedBytes: Uint8Array | null = null
  const approvals = new Map<string, string>(), plans = new Map<string, AgentChangeSetEnvelope>(), previews = new Map<string, NativePptxDeck>(), proposalRevisions = new Map<string, string>()
  const traces: Array<{ request: AgentToolCall; response: AgentToolCallResult }> = []
  const proposalKey = (operations: readonly AgentDemoOperation[]) => JSON.stringify(operations.map((item) => [item.op, item.input]))
  const assertOpen = () => { if (disposed) throw new Error('This presentation session is closed. Reload the sample to continue.') }
  const dispose = () => { if (!disposed) { disposed = true; verifiedBytes = null; approvals.clear(); previews.clear(); runtime.terminate() } }
  try {
    const response = await fetcher(SAMPLE_PATH)
    if (!response.ok) throw new Error(`The sample presentation could not be loaded (HTTP ${response.status}).`)
    let bytes = new Uint8Array(await response.arrayBuffer())
    let deck = (await runtime.extract(bytes.slice())).deck
    assertNativePptx(deck)
    if (deck.origin !== 'parsed' || !deck.sourceRevision) throw new Error('The presentation must be an exact parsed source package.')
    let selectedOperation: AgentDemoOperation | undefined
    const artifactView = (value = deck): AgentDemoArtifact => ({ artifactId: 'northstar-agent-presentation', format: 'pptx', name: AGENT_PPTX_NAME, revision: value.sourceRevision!, content: projection(value, selectedOperation) })
    let queue = Promise.resolve()
    const serialize = <T>(action: () => Promise<T>): Promise<T> => { const result = queue.then(action); queue = result.then(() => undefined, () => undefined); return result }
    const ledger = new Map<string, { signature: string; deck: NativePptxDeck }>()
    const applyCopy = async (request: PptxNativeMutationRequest) => {
      assertOpen()
      const source = deck
      if (source.sourceRevision !== request.expectedSourceRevision) throw new Error('The source revision changed before the native write.')
      const applied = await runtime.apply({ original: bytes.slice(), deck: structuredClone(source), mutation: structuredClone(request), sourceName: AGENT_PPTX_NAME })
      const replacement = applied.bytes.slice()
      const reopened = (await runtime.extract(replacement.slice())).deck
      assertNativePptx(reopened)
      if (reopened.origin !== 'parsed' || !reopened.sourceRevision || reopened.sourceRevision === source.sourceRevision) throw new Error('Native replacement did not produce a new exact source revision.')
      assertOpen()
      return { bytes: replacement, deck: reopened, sourceRevision: source.sourceRevision }
    }
    const host: NativePptxAgentArtifact = {
      kind: 'injoffice.pptx-native', artifactId: 'northstar-agent-presentation',
      async snapshot() {
        assertOpen()
        if (postCommitReads !== null && postCommitReads++ === 1) { postCommitReads = null; if (failVerification) throw new Error('Demo fault: post-commit verification read unavailable. The write already succeeded.') }
        const current = (await runtime.extract(bytes.slice())).deck
        assertOpen()
        return { deck: current }
      },
      async preview(request) { const candidate = await applyCopy(request); previews.set(candidate.deck.sourceRevision!, candidate.deck); return { deck: candidate.deck } },
      apply(request, options) { return serialize(async () => {
        assertOpen()
        const signature = JSON.stringify(request), retained = ledger.get(options.idempotencyKey)
        if (retained) { if (retained.signature !== signature) throw new Error('A different native mutation used this idempotency key.'); return { snapshot: { deck: structuredClone(retained.deck) }, deduplicated: true } }
        const candidate = await applyCopy(request)
        if (deck.sourceRevision !== candidate.sourceRevision) throw new Error('The source revision changed during the native write.')
        bytes = candidate.bytes; deck = candidate.deck; nativeWrites++; verifiedBytes = null; postCommitReads = 0
        ledger.set(options.idempotencyKey, { signature, deck: structuredClone(deck) })
        return { snapshot: { deck: structuredClone(deck) }, receipt: { nativeWrites }, deduplicated: false }
      }) },
    }
    const core = await createAgentSession({ artifact: host, adapter: createNativePptxAgentAdapter(), actor: { id: 'playground-local-host', kind: 'agent' },
      confirmDestructive: ({ changeSet, identity }) => !!changeSet && approvals.get(changeSet.changeSetId) === JSON.stringify(changeSet) && changeSet.baseRevision === identity.revision && changeSet.baseFingerprint === identity.fingerprint })
    const dispatcher = createAgentToolDispatcher(core)
    let requestNumber = 0
    const call = async <T>(method: AgentToolMethod, params: JsonObject = {}): Promise<T> => {
      assertOpen()
      const request: AgentToolCall = { protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId: `pptx-tool-${++requestNumber}`, method, params }
      const result = await dispatcher.dispatch(request)
      assertOpen(); traces.push(structuredClone({ request, response: result }))
      if (!result.ok) throw new Error(`${result.error.code}: ${result.error.message}${result.error.issues?.map((item) => ` ${item.code}: ${item.message}`).join('') ?? ''}`)
      return result.result as T
    }
    let disclosed: { revision: string; targets: MockPptxTarget[] } | undefined
    const proposalContext = async (): Promise<JsonObject> => {
      const capabilities = await call<{ capabilities: AgentCapability[] }>('office.capabilities')
      const inspection = await call<AgentBoundedResult>('office.inspect', BOUNDS)
      if (inspection.truncated) throw new Error('The presentation structure exceeds the bounded proposal context.')
      const targets: MockPptxTarget[] = []
      let cursor: string | undefined, count = 0
      const slideByElement = new Map(elements(deck).map(({ slideId, element }) => [element.id, slideId]))
      do {
        const result = await call<AgentBoundedResult>('office.read', { ...BOUNDS, ...(cursor ? { cursor } : {}) })
        if (result.identity.revision !== inspection.identity.revision) throw new Error('The document revision changed during proposal discovery.')
        const read = (result.data as unknown as { elements: NativeElement[] }).elements
        count += read.length
        if (count > 1000 || result.truncated && !result.nextCursor) throw new Error('The presentation exceeds the bounded element discovery budget.')
        for (const element of read) if (exactTarget(element)) targets.push({ elementId: element.id, currentText: firstText(element), slideId: slideByElement.get(element.id)! })
        cursor = result.nextCursor
      } while (cursor)
      if (targets.length > 256) throw new Error('The presentation exceeds the bounded editable target budget.')
      disclosed = { revision: inspection.identity.revision, targets }
      return { format: 'pptx', identity: { artifactId: inspection.identity.artifactId, revision: inspection.identity.revision, fingerprint: inspection.identity.fingerprint },
        capabilities: capabilities.capabilities.filter((item) => item.name === 'pptx.native.text.replace') as unknown as JsonObject[], constraints: { maxOperations: 1, allowedTargets: targets.map((target) => ({ ...target })) } }
    }
    const acceptProposal = (value: unknown): AgentDemoOperation[] => {
      assertOpen()
      if (!disclosed) throw new Error('Read and disclose the presentation context before accepting a proposal.')
      if (!object(value) || Object.keys(value).some((key) => key !== 'operations') || !Array.isArray(value.operations) || value.operations.length !== 1) throw new Error('A proposal must contain exactly one operation and no other fields.')
      const operation = value.operations[0]
      if (!object(operation) || Object.keys(operation).some((key) => !['name', 'input', 'operationId'].includes(key)) || operation.name !== 'pptx.native.text.replace' || !object(operation.input)) throw new Error('Only a bounded native presentation text replacement is accepted.')
      const input = operation.input
      if (Object.keys(input).some((key) => !['elementId', 'text'].includes(key)) || typeof input.text !== 'string' || !input.text || input.text.length > 1000 || /[\t\r\n]/.test(input.text)) throw new Error('The proposal requires an exact element id and bounded single-line text.')
      const matches = disclosed.targets.filter((target) => target.elementId === input.elementId)
      if (matches.length !== 1 || matches[0].currentText === input.text) throw new Error('The proposal must change exactly one disclosed editable target.')
      if (operation.operationId !== undefined && (typeof operation.operationId !== 'string' || !/^[A-Za-z0-9._:-]{1,128}$/.test(operation.operationId))) throw new Error('The operation id must be a short identifier.')
      const accepted = [{ operationId: typeof operation.operationId === 'string' ? operation.operationId : 'agent-pptx-text', op: operation.name, target: matches[0].elementId, value: input.text, input: { elementId: matches[0].elementId, text: input.text } }]
      selectedOperation = accepted[0]; proposalRevisions.set(proposalKey(accepted), disclosed.revision)
      return accepted
    }
    const propose = async (request: string): Promise<AgentDemoOperation[]> => {
      const context = await proposalContext()
      if (/\bmacro\b/i.test(request)) return [{ operationId: 'agent-pptx-refusal', op: 'macro.execute', target: 'presentation', input: {} }]
      const { capabilities, ...bounded } = context
      return acceptProposal(createMockPptxProposal({ request, context: bounded, capabilities }))
    }
    const session: AgentDemoSession = {
      async capabilities() { const result = await call<{ capabilities: AgentCapability[] }>('office.capabilities'); return { format: 'pptx', operations: result.capabilities.map((item) => ({ operation: item.name, access: 'write', destructive: item.destructive })) } },
      async inspect({ selection, maxItems, maxBytes }) { const result = await call<AgentBoundedResult>('office.inspect', { maxItems, maxBytes }); return { artifactId: result.identity.artifactId, revision: result.identity.revision, fingerprint: result.identity.fingerprint, selection, content: projection(deck, selectedOperation), truncated: result.truncated } },
      async plan(operations, { expectedRevision }) {
        const proposalRevision = proposalRevisions.get(proposalKey(operations))
        if (proposalRevision && proposalRevision !== expectedRevision) throw new Error('The document revision changed after proposal discovery. Prepare a new proposal.')
        const envelope = await call<AgentChangeSetEnvelope>('office.plan', { expectedRevision, operations: operations.map((item) => ({ operationId: item.operationId, name: item.op, input: item.input as JsonObject })) })
        plans.set(envelope.changeSetId, envelope)
        const params = { changeSetId: envelope.changeSetId }, before = structuredClone(deck), selected = operations[0]
        return { id: envelope.changeSetId, expectedRevision: envelope.baseRevision, operations: structuredClone(operations),
          async validate() { const result = await call<AgentValidationReport>('office.validate', params); return { ok: result.valid, issues: result.issues.map((item) => ({ code: item.code.toLowerCase(), path: item.path ?? '', message: item.message })) } },
          async preview() { const result = await call<AgentArtifactView>('office.preview', params); const errors = result.issues.filter((item) => item.severity === 'error'); if (errors.length) throw new Error(`Native preview failed: ${errors.map((item) => `${item.code}: ${item.message}`).join('; ')}. No source write was committed.`); const candidate = previews.get(String((result.data as JsonObject).resultingRevision)) ?? before; return { artifact: { ...artifactView(candidate), content: projection(candidate, selected) }, summary: result.issues.length ? result.issues.map((item) => item.message).join('; ') : 'Native PPTX preview applied to a copy and reopened; source unchanged.', evidence: result.evidence.map((item) => item.description ?? item.kind) } },
          async diff() { const result = await call<AgentArtifactView>('office.diff', params); return { changes: result.issues.length ? [] : operations.map((operation) => ({ target: operation.target, before: firstText(elements(before).find(({ element }) => element.id === operation.input.elementId)?.element), after: String(operation.value ?? '') })) } },
          async commit(input) { if (input.expectedRevision !== envelope.baseRevision) throw new Error('Commit revision does not match the reviewed change set.'); const result = await call<AgentCommitResult>('office.commit', { ...params, idempotencyKey: input.idempotencyKey, confirmation: input.confirmation }); verifiedBytes = result.verification.verified && result.identity.revision === deck.sourceRevision && result.identity.fingerprint === deck.sourceRevision ? bytes.slice() : null; return { artifactId: result.identity.artifactId, previousRevision: envelope.baseRevision, revision: result.identity.revision, fingerprint: result.identity.fingerprint, operationIds: envelope.operations.map((item) => item.operationId), evidence: result.evidence?.map((item) => item.description ?? item.kind) ?? [], verification: result.verification, content: projection(deck, selected) } },
          async verify(receipt) { return { ok: receipt.verification.verified, revision: receipt.revision, fingerprint: receipt.fingerprint, evidence: receipt.verification.checks.map((item) => item.message ?? item.name) } },
        }
      },
    }
    await proposalContext()
    const first = disclosed!.targets.find((target, _, targets) => targets.filter((other) => other.currentText === target.currentText).length === 1)
    if (!first) throw new Error('The sample has no unique exact text target.')
    const prompt = mode === 'safe' ? `Replace "${first.currentText}" with "Northstar: ready for launch"` : 'Execute an unsupported presentation macro.'
    const operations = await propose(prompt)
    const scenario: AgentDemoScenario = { artifact: artifactView(), operations, prompt, summary: mode === 'safe' ? 'Discover exact slide text, review a native preview, approve the change, and verify the saved PPTX.' : 'Refuse macro execution before any presentation write.' }
    return { scenario, operations, session, dispose, propose, proposalContext, acceptProposal,
      async approve(changeSetId: string) { assertOpen(); const envelope = plans.get(changeSetId); if (!envelope) throw new Error('Only an exact change set retained by this host can be approved.'); const current = await call<AgentBoundedResult>('office.inspect', BOUNDS); if (current.identity.revision !== envelope.baseRevision || current.identity.fingerprint !== envelope.baseFingerprint) throw new Error('The document revision changed. Prepare and review a new proposal.'); approvals.set(changeSetId, JSON.stringify(envelope)) },
      async simulateConcurrentEdit() { await serialize(async () => {
        assertOpen()
        const target = elements(deck).map((item) => item.element).find(exactTarget)
        if (!target || target.kind !== 'text' && target.kind !== 'shape') throw new Error('No exact text target is available for the concurrent edit.')
        const paragraphs = structuredClone(target.paragraphs) as PptxNativeExactParagraph[]
        paragraphs[0].runs[0].text += ' (reviewed)'
        const candidate = await applyCopy({ expectedSourceRevision: deck.sourceRevision!, operations: [{ operationId: 'host-concurrent-edit', kind: 'text.replace', elementId: target.id, expectedFingerprintSha256: target.source!.fingerprintSha256, paragraphs }] })
        bytes = candidate.bytes; deck = candidate.deck; nativeWrites++; verifiedBytes = null; approvals.clear(); postCommitReads = null
      }) },
      setVerificationFailure(enabled: boolean) { assertOpen(); failVerification = enabled }, trace: () => structuredClone(traces), stats: () => ({ nativeWrites }),
      download: () => !disposed && verifiedBytes ? new Blob([new Uint8Array(verifiedBytes).buffer], { type: MEDIA_TYPE }) : null,
    }
  } catch (error) { dispose(); throw error }
}
