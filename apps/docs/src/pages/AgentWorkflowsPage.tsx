import { Article } from '../components/Article'
import { Callout } from '../components/Callout'
import { CodeBlock } from '../components/CodeBlock'
import { Install } from '../components/Install'

export function AgentWorkflowsPage() {
  return (
    <Article id="agent-workflows">
      <h2 id="architecture">A document engine for any agent</h2>
      <p>InjOffice exposes document work as data and leaves model choice, prompts, chat UI, and orchestration to your application. The agent proposes operations. InjOffice owns validation, isolated preview, revision checks, commit, and verification.</p>
      <CodeBlock language="bash" title="execution boundary" code={`model or human request
        ↓ proposes plain operations
AgentSession → inspect → plan → preview / diff → validate
                                                ↓
                                      host approval policy
                                                ↓
                                      commit → reopen → verify`} />
      <Callout kind="note" title="Provider independent">The packages do not contain a model SDK and do not call an InjOffice service. Use OpenAI, Anthropic, a local model, deterministic rules, or a human-authored operation list.</Callout>

      <h2 id="quickstart">Create a guarded change set</h2>
      <Install packages="@injoffice/agent-tools @injoffice/agent-office" />
      <CodeBlock language="ts" code={`import { createAgentSession } from '@injoffice/agent-tools'
import { createXlsxAgentAdapter } from '@injoffice/agent-office/xlsx'

const session = await createAgentSession({
  artifact, // The host-owned artifact passed to the XLSX adapter.
  adapter: createXlsxAgentAdapter(),
  actor: { id: 'forecast-agent', kind: 'agent', displayName: 'Forecast agent' },
  limits: { maxReadItems: 500, maxReadBytes: 64_000, maxOperations: 100, maxOperationBytes: 32_000 },
  confirmDestructive: async ({ confirmation }) => approvalStore.accepts(confirmation),
})

const report = await session.capabilities()
const context = await session.inspect({ query: { selection: 'Forecast!A1:D20' }, maxItems: 100, maxBytes: 16_000 })

// Your model produces only this plain operation proposal.
const changeSet = await session.plan([
  { name: 'xlsx.cell.set_value', operationId: 'agent-op-1', input: {
    sheetId: 'sheet-forecast', cell: { row: 4, column: 3 }, value: 'High',
  } },
], {
  expectedRevision: context.identity.revision,
  expectedFingerprint: context.identity.fingerprint,
})

const [preview, diff, validation, plannedProof] = await Promise.all([
  changeSet.preview(), changeSet.diff(), changeSet.validate(), changeSet.verify(),
])
if (!validation.valid || !plannedProof.verified) throw new Error(validation.issues[0]?.message ?? 'Verification failed')

const review = await showReviewUI({ preview, diff })
const receipt = await changeSet.commit({
  idempotencyKey: crypto.randomUUID(),
  confirmation: { approvalId: review.approvalId },
})
if (!receipt.verification.verified) throw new Error('Committed output verification failed')`} />

      <h2 id="capabilities">Discover capabilities at runtime</h2>
      <p>Do not teach an agent a fictional universal Office API. Ask the selected adapter what it supports for this artifact and version. Capability records distinguish reads, writes, destructive operations, and required confirmation.</p>
      <ul>
        <li><strong>XLSX:</strong> bounded inspection of native v1/v2 projections plus v1 cell, formula, style, dimension, and merge mutations.</li>
        <li><strong>DOCX:</strong> bounded document inspection and semantic guarded text replacement.</li>
        <li><strong>PPTX:</strong> authored-deck slide operations plus source-anchored native text and AutoShape updates.</li>
        <li><strong>PDF:</strong> bounded page inspection and rotate, insert, delete, reorder, crop, resize, or n-up operations.</li>
      </ul>
      <p>Support remains operation-specific. Encrypted files, macros, embedded objects, ambiguous targets, or features outside an adapter’s contract are refused rather than approximated.</p>

      <h2 id="safety">Safety is part of the API</h2>
      <ul>
        <li>Inspections are bounded and identify the exact artifact revision and fingerprint.</li>
        <li>Plans are immutable, operation-limited, and separate from the source artifact.</li>
        <li>Preview and semantic diff happen before write.</li>
        <li>Validation returns stable issue codes and JSON paths an agent can repair.</li>
        <li>Commit uses compare-and-swap revisions and idempotency keys.</li>
        <li>Host policy controls approval for destructive or high-impact changes.</li>
        <li>Verification reopens committed output and returns machine-readable evidence.</li>
      </ul>
      <Callout kind="caution">A preview is not a commit, and a successful commit is not verification. Preserve these boundaries in your agent UI and audit trail.</Callout>

      <h2 id="non-goals">What the package deliberately does not own</h2>
      <ul>
        <li>Model hosting, API keys, provider routing, or token billing</li>
        <li>A proprietary chat panel or autonomous agent loop</li>
        <li>Web search, image generation, or connector credentials</li>
        <li>Your authorization rules, tenant isolation, or durable artifact storage</li>
        <li>A promise that every Office feature can be edited</li>
      </ul>
      <p>This boundary keeps InjOffice usable as an open-source library inside a browser app, local tool, server, MCP adapter, or a product-specific agent experience.</p>

      <h2 id="demo">Run the local proof</h2>
      <p>Open <a href="#/agent">Agent workflows</a> in the workbench. The demo supplies a deterministic local proposal so you can inspect the execution contract without an API key or network request. Switch among XLSX, DOCX, PPTX, and PDF, then run the refusal proof to see an unsupported request stop before write.</p>
    </Article>
  )
}
