import { createPdfAgentAdapter, type PdfAgentArtifact } from '@injoffice/agent-office/pdf'
import { applyPageOps, readInfo, normalizeRotation, type PdfDocumentInfo } from '@injoffice/pdf/browser'
import { AGENT_TOOLS_PROTOCOL, AGENT_TOOLS_PROTOCOL_VERSION, createAgentSession, createAgentToolDispatcher,
  type AgentArtifactAdapter, type AgentArtifactView, type AgentBoundedResult, type AgentCapability, type AgentChangeSetEnvelope,
  type AgentCommitResult, type AgentToolCall, type AgentToolCallResult, type AgentToolMethod, type AgentValidationReport, type JsonObject } from '@injoffice/agent-tools'
import type { AgentDemoSession } from './agentDemoRuntime'
import type { AgentDemoArtifact, AgentDemoMode, AgentDemoOperation, AgentDemoScenario } from './agentDemoScenario'
import { createPdfDemoFixture, PDF_DEMO_FILE_NAME } from './pdfDemoFixture'
import { createMockPdfProposal } from './mockPdfProposal'

const BOUNDS = { maxItems: 100, maxBytes: 32_000 }
const DEGREES = [90, 180, 270, -90, -180, -270]
const object = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value)
const key = (operations: readonly AgentDemoOperation[]) => JSON.stringify(operations.map((item) => [item.op, item.input]))
function rotatedInfo(before: PdfDocumentInfo, operations: readonly { name: string; input: Record<string, unknown> }[]): PdfDocumentInfo {
  const expected = structuredClone(before)
  for (const operation of operations) {
    if (operation.name !== 'pdf.page.rotate' || !Array.isArray(operation.input.pages)) throw new Error('The demo only persists PDF rotation proposals.')
    for (const page of expected.pages) if (operation.input.pages.includes(page.index)) page.rotation = normalizeRotation(page.rotation + Number(operation.input.degrees))
  }
  return expected
}
function sameInfo(expected: PdfDocumentInfo, candidate: unknown): candidate is PdfDocumentInfo {
  if (!object(candidate) || candidate.pageCount !== expected.pageCount || !Array.isArray(candidate.pages) || candidate.pages.length !== expected.pages.length) return false
  return candidate.pages.every((page, index) => object(page) && ['index', 'width', 'height', 'rotation'].every((field) => page[field] === expected.pages[index][field as keyof PdfDocumentInfo['pages'][number]]))
}

/** Browser-owned real PDF bytes, using the shipped page adapter and public dispatcher. */
export async function createNativePdfAgentSessionInput(mode: AgentDemoMode, fixture = createPdfDemoFixture) {
  const host: { kind: 'injoffice.pdf'; artifactId: string; bytes: Uint8Array } = { kind: 'injoffice.pdf', artifactId: 'northstar-agent-pdf', bytes: (await fixture()).slice() }
  let info = await readInfo(host.bytes)
  let disposed = false; let nativeWrites = 0; let failVerification = false
  let selectedPage = 2; let verifiedBytes: Uint8Array | null = null
  let disclosed: { revision: string; pages: number[] } | undefined
  const approvals = new Map<string, string>(); const plans = new Map<string, AgentChangeSetEnvelope>()
  const proposalRevisions = new Map<string, string>()
  const expectedOutputs = new Map<string, PdfDocumentInfo>()
  const traces: Array<{ request: AgentToolCall; response: AgentToolCallResult }> = []
  const assertOpen = () => { if (disposed) throw new Error('This PDF session is closed. Reload the sample to continue.') }
  let queue = Promise.resolve()
  const serialize = <T>(action: () => Promise<T>): Promise<T> => { const result = queue.then(action); queue = result.then(() => undefined, () => undefined); return result }
  const projection = (value: PdfDocumentInfo, page = selectedPage): Record<string, unknown> => ({ title: 'Northstar operating review', pageCount: value.pageCount,
    pages: value.pages.map((entry) => ({ page: entry.index, rotation: entry.rotation, width: entry.width, height: entry.height })),
    selectedPage: page, rotation: value.pages.find((entry) => entry.index === page)?.rotation ?? 0 })
  const publicAdapter = createPdfAgentAdapter()
  const adapter: AgentArtifactAdapter<PdfAgentArtifact> = {
    ...publicAdapter,
    // This demo host requires human approval even for non-destructive rotations.
    async capabilities(request) { return (await publicAdapter.capabilities(request)).map((item) => ({ ...item, requiresConfirmation: true })) },
    commit(request) { return serialize(async () => {
      assertOpen()
      const before = await readInfo(host.bytes)
      const result = await publicAdapter.commit(request)
      assertOpen()
      const expected = rotatedInfo(before, request.changeSet.operations)
      host.bytes = result.artifact.bytes.slice(); info = await readInfo(host.bytes); nativeWrites += 1; verifiedBytes = null
      expectedOutputs.set(request.changeSet.changeSetId, expected)
      return { ...result, artifact: host }
    }) },
    async verify(request) {
      const result = await publicAdapter.verify(request)
      if (request.stage === 'planned') return result
      const reopened = await readInfo(request.artifact.bytes)
      const expected = expectedOutputs.get(request.changeSet.changeSetId)
      const checks = [...result.checks,
        { name: 'pdf-exact-output-page-rotation-and-geometry', passed: !!expected && sameInfo(expected, reopened), message: 'Reopened output has the requested rotation, unchanged page count, and unchanged page geometry.' },
        { name: 'pdf-demo-verification-read', passed: !failVerification, message: failVerification ? 'Demo fault: write completed, but verification was intentionally failed.' : 'Verification fault injection is disabled.' }]
      return { ...result, checks, verified: result.verified && checks.every((check) => check.passed) }
    },
  }
  const core = await createAgentSession({ artifact: host, adapter, actor: { id: 'playground-pdf-host', kind: 'agent' },
    confirmDestructive: ({ changeSet, identity }) => !!changeSet && approvals.get(changeSet.changeSetId) === JSON.stringify(changeSet) && changeSet.baseRevision === identity.revision && changeSet.baseFingerprint === identity.fingerprint })
  const dispatcher = createAgentToolDispatcher(core)
  let requestNumber = 0
  const call = async <T>(method: AgentToolMethod, params: JsonObject = {}): Promise<T> => {
    assertOpen()
    const request: AgentToolCall = { protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION, requestId: `pdf-demo-tool-${++requestNumber}`, method, params }
    const result = await dispatcher.dispatch(request); assertOpen(); traces.push(structuredClone({ request, response: result }))
    if (!result.ok) throw new Error(`${result.error.code}: ${result.error.message}${result.error.issues?.length ? ` ${result.error.issues.map((issue) => `${issue.code}: ${issue.message}`).join('; ')}` : ''}`)
    return result.result as T
  }
  const artifactView = (value: PdfDocumentInfo, revision: string, page = selectedPage): AgentDemoArtifact => ({ artifactId: host.artifactId, format: 'pdf', name: PDF_DEMO_FILE_NAME, revision, content: projection(value, page) })
  const proposalContext = async (): Promise<JsonObject> => {
    const capabilities = await call<{ capabilities: AgentCapability[] }>('office.capabilities')
    const inspection = await call<AgentBoundedResult>('office.inspect', BOUNDS)
    const read = await call<AgentBoundedResult>('office.read', BOUNDS)
    if (inspection.truncated || read.truncated || inspection.identity.revision !== read.identity.revision) throw new Error('PDF discovery was truncated or its source revision changed.')
    const pages = (read.data as unknown as { pages: PdfDocumentInfo['pages'] }).pages
    disclosed = { revision: read.identity.revision, pages: pages.map((page) => page.index) }
    return { format: 'pdf', identity: { ...read.identity }, document: { title: 'Northstar operating review', pages: pages.map((page) => ({ ...page })) },
      capabilities: capabilities.capabilities.filter((item) => item.name === 'pdf.page.rotate') as unknown as JsonObject[],
      constraints: { maxOperations: 1, allowedPages: disclosed.pages, allowedDegrees: DEGREES } }
  }
  const acceptProposal = (value: unknown): AgentDemoOperation[] => {
    assertOpen()
    if (!disclosed) throw new Error('Read PDF proposal context before accepting a proposal.')
    if (!object(value) || Object.keys(value).some((field) => field !== 'operations') || !Array.isArray(value.operations) || value.operations.length !== 1) throw new Error('Only one PDF operation is accepted.')
    const operation = value.operations[0]
    if (!object(operation) || Object.keys(operation).some((field) => !['name', 'input', 'operationId'].includes(field)) || operation.name !== 'pdf.page.rotate' || !object(operation.input)) throw new Error('Only a bounded PDF page rotation is accepted.')
    const input = operation.input
    if (Object.keys(input).some((field) => !['pages', 'degrees'].includes(field)) || !Array.isArray(input.pages) || input.pages.length !== 1 || !disclosed.pages.includes(input.pages[0]) || typeof input.degrees !== 'number' || !DEGREES.includes(input.degrees)) throw new Error('The PDF rotation must use one disclosed page and an allowed degree increment.')
    if (operation.operationId !== undefined && (typeof operation.operationId !== 'string' || !/^[A-Za-z0-9._:-]{1,128}$/.test(operation.operationId))) throw new Error('Invalid operation identifier.')
    selectedPage = input.pages[0]
    const accepted = [{ operationId: typeof operation.operationId === 'string' ? operation.operationId : `pdf-rotate-${selectedPage}`, op: 'pdf.page.rotate', target: `page-${selectedPage}`, value: input.degrees, input: { pages: [selectedPage], degrees: input.degrees } }]
    proposalRevisions.set(key(accepted), disclosed.revision)
    return accepted
  }
  const propose = async (request: string) => {
    const { capabilities, ...context } = await proposalContext()
    return acceptProposal(createMockPdfProposal({ request, context, capabilities }))
  }
  const session: AgentDemoSession = {
    async capabilities() { const result = await call<{ capabilities: AgentCapability[] }>('office.capabilities'); return { format: 'pdf', operations: result.capabilities.map((item) => ({ operation: item.name, access: 'write', destructive: item.destructive })) } },
    async inspect({ selection, maxItems, maxBytes }) { const result = await call<AgentBoundedResult>('office.inspect', { maxItems, maxBytes }); return { artifactId: result.identity.artifactId, revision: result.identity.revision, fingerprint: result.identity.fingerprint, selection, content: projection(info), truncated: result.truncated } },
    async plan(operations, { expectedRevision }) {
      if (proposalRevisions.has(key(operations)) && proposalRevisions.get(key(operations)) !== expectedRevision) throw new Error('The PDF source revision changed after discovery. Prepare a new proposal.')
      const envelope = await call<AgentChangeSetEnvelope>('office.plan', { expectedRevision, operations: operations.map((item) => ({ operationId: item.operationId, name: item.op, input: item.input as JsonObject })) })
      plans.set(envelope.changeSetId, envelope)
      const params = { changeSetId: envelope.changeSetId }; const before = structuredClone(info); const page = selectedPage
      return { id: envelope.changeSetId, expectedRevision: envelope.baseRevision, operations: structuredClone(operations),
        async validate() { const result = await call<AgentValidationReport>('office.validate', params); return { ok: result.valid, issues: result.issues.map((item) => ({ code: item.code.toLowerCase(), path: item.path ?? '', message: item.message })) } },
        async preview() {
          const result = await call<AgentArtifactView>('office.preview', params)
          const errors = result.issues.filter((issue) => issue.severity === 'error')
          if (errors.length) throw new Error(`PDF preview failed: ${errors.map((issue) => `${issue.code}: ${issue.message}`).join('; ')}. No source write was committed.`)
          const refused = result.issues.some((issue) => issue.severity === 'refusal') && object(result.data) && result.data.mode === 'refused'
          if (refused) return { artifact: artifactView(before, envelope.baseRevision, page), summary: 'Unsupported PDF operation; no source file was written.', evidence: result.evidence.map((item) => item.description ?? item.kind) }
          const expected = rotatedInfo(before, envelope.operations)
          if (!sameInfo(expected, result.data)) throw new Error('PDF preview failed: PREVIEW_MISMATCH. Reopened preview must prove the requested rotation, page count, and unchanged page geometry. No source write was committed.')
          return { artifact: artifactView(result.data, envelope.baseRevision, page), summary: 'Real PDF rotation preview; the source bytes are unchanged.', evidence: result.evidence.map((item) => item.description ?? item.kind) }
        },
        async diff() {
          const validation = await call<AgentValidationReport>('office.validate', params)
          if (!validation.valid) return { changes: [] }
          const result = await call<AgentArtifactView>('office.diff', params)
          const data = result.data as unknown as { before: PdfDocumentInfo; after: PdfDocumentInfo }
          return { changes: data.after.pages.filter((entry, index) => entry.rotation !== data.before.pages[index]?.rotation).map((entry) => ({ target: `page-${entry.index}`, before: `${data.before.pages[entry.index - 1].rotation}°`, after: `${entry.rotation}°` })) }
        },
        async commit(input) {
          if (input.expectedRevision !== envelope.baseRevision) throw new Error('Commit revision does not match the reviewed PDF plan.')
          const result = await call<AgentCommitResult>('office.commit', { ...params, idempotencyKey: input.idempotencyKey, confirmation: input.confirmation })
          const current = await publicAdapter.identity({ artifact: host, actor: { id: 'playground-pdf-host', kind: 'agent' } })
          verifiedBytes = result.verification.verified && result.identity.fingerprint === current.fingerprint ? host.bytes.slice() : null
          return { artifactId: result.identity.artifactId, previousRevision: envelope.baseRevision, revision: result.identity.revision, fingerprint: result.identity.fingerprint,
            operationIds: envelope.operations.map((item) => item.operationId), evidence: result.evidence?.map((item) => item.description ?? item.kind) ?? [], verification: result.verification, content: projection(info, page) }
        },
        async verify(receipt) { return { ok: receipt.verification.verified, revision: receipt.revision, fingerprint: receipt.fingerprint, evidence: receipt.verification.checks.map((check) => check.message ?? check.name) } },
      }
    },
  }
  const prompt = mode === 'safe' ? 'Rotate page 2 by 90 degrees' : 'Rewrite the PDF text using an unsupported operation'
  const operations: AgentDemoOperation[] = mode === 'safe' ? await propose(prompt) : [{ operationId: 'pdf-refusal', op: 'pdf.text.rewrite', target: 'page-2', input: {} }]
  const scenario: AgentDemoScenario = { artifact: artifactView(info, core.identity.revision), prompt, summary: mode === 'safe' ? 'Rotate one real PDF page while preserving the document and page geometry.' : 'Unsupported PDF operation; no output file is written.', operations }
  return { scenario, operations, session, propose, proposalContext, acceptProposal,
    async approve(changeSetId: string) { assertOpen(); const plan = plans.get(changeSetId); if (!plan) throw new Error('Unknown PDF plan.'); const inspection = await call<AgentBoundedResult>('office.inspect', BOUNDS); if (inspection.identity.revision !== plan.baseRevision || inspection.identity.fingerprint !== plan.baseFingerprint) throw new Error('The PDF source revision changed after review.'); approvals.set(changeSetId, JSON.stringify(plan)) },
    simulateConcurrentEdit() { return serialize(async () => { assertOpen(); const replacement = await applyPageOps(host.bytes, [{ type: 'rotate', pages: [1], degrees: 90 }]); assertOpen(); host.bytes = replacement; info = await readInfo(host.bytes); nativeWrites += 1; verifiedBytes = null; approvals.clear() }) },
    setVerificationFailure(enabled: boolean) { assertOpen(); failVerification = enabled },
    trace: () => structuredClone(traces), stats: () => ({ nativeWrites }),
    download: () => !disposed && verifiedBytes ? new Blob([new Uint8Array(verifiedBytes).buffer], { type: 'application/pdf' }) : null,
    dispose() { disposed = true; verifiedBytes = null; approvals.clear(); plans.clear(); proposalRevisions.clear(); expectedOutputs.clear(); disclosed = undefined },
  }
}
