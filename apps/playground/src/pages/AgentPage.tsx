import { useEffect, useMemo, useState } from 'react'
import {
  createDemoSessionInput,
  type AgentCapability,
  type AgentCommitReceipt,
  type AgentDemoChangeSet,
  type AgentDiff,
  type AgentInspection,
  type AgentPreview,
  type AgentValidation,
  type AgentVerification,
} from '../agentDemoRuntime'
import { createAgentDemoScenario, type AgentDemoFormat, type AgentDemoMode } from '../agentDemoScenario'
import { AGENT_TOOLS, agentFormatFromTool, agentHref, parseAgentTool, type AgentTool } from '../route'
import { DsButton, DsCallout, DsChip, DsSegment } from '../design-system/primitives'
import '../design-system/live-tools.css'

type WorkflowState = 'ready' | 'preparing' | 'awaiting-approval' | 'refused' | 'committing' | 'verified' | 'error'
type WorkflowStep = 'Inspect' | 'Plan' | 'Preview + diff' | 'Validate' | 'Approve' | 'Commit' | 'Verify'
type AgentToolName = 'office.capabilities' | 'office.inspect' | 'office.plan' | 'office.preview' | 'office.diff' | 'office.validate' | 'office.commit' | 'office.verify'

const STEPS: WorkflowStep[] = ['Inspect', 'Plan', 'Preview + diff', 'Validate', 'Approve', 'Commit', 'Verify']
const AGENT_TOOL_METHODS: AgentToolName[] = ['office.capabilities', 'office.inspect', 'office.plan', 'office.preview', 'office.diff', 'office.validate', 'office.commit', 'office.verify']

function stepState(step: WorkflowStep, state: WorkflowState): 'done' | 'active' | 'waiting' | 'refused' {
  const progress: Record<WorkflowState, number> = {
    ready: 0,
    preparing: 2,
    'awaiting-approval': 4,
    refused: 3,
    committing: 5,
    verified: 7,
    error: 0,
  }
  const index = STEPS.indexOf(step)
  if (state === 'refused' && step === 'Validate') return 'refused'
  if (index < progress[state]) return 'done'
  if (index === progress[state]) return 'active'
  return 'waiting'
}

function SheetArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const headers = content.headers as string[]
  const rows = content.rows as Array<Array<string | number>>
  return (
    <table className="agent-artifact__sheet ds-table">
      <thead><tr><th aria-label="Row number" />{headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
      <tbody>{rows.map((row, rowIndex) => <tr key={String(row[0])}><th>{rowIndex + 2}</th>{row.map((cell, cellIndex) => <td className={highlighted && rowIndex === 3 && cellIndex === 3 ? 'agent-artifact__changed' : undefined} key={`${rowIndex}-${cellIndex}`}>{cell}</td>)}</tr>)}</tbody>
    </table>
  )
}

function DocxArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const blocks = content.blocks as Array<{ id: string; kind: string; text: string }>
  return (
    <article className="agent-artifact__doc ds-page">
      <h3>{String(content.title)}</h3>
      {blocks.map((block) => <p className={highlighted && block.id === 'block-summary' ? 'agent-artifact__changed' : undefined} key={block.id}>{block.text}</p>)}
    </article>
  )
}

function PptxArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const metric = content.metric as { label: string; value: string }
  return (
    <article className="agent-artifact__slide ds-slide">
      <span>Board update</span>
      <h3>{String(content.title)}</h3>
      <p>{String(content.subtitle)}</p>
      <div className={`agent-artifact__metric${highlighted ? ' agent-artifact__changed' : ''}`}><strong>{metric.value}</strong><small>{metric.label}</small></div>
    </article>
  )
}

function PdfArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  return (
    <article className="agent-artifact__pdf ds-pdf-sheet">
      <span>Review packet · page 2 of {String(content.pageCount)}</span>
      <h3>{String(content.heading)}</h3>
      <p>{String(content.body)}</p>
      <aside className={`agent-artifact__note${highlighted ? ' agent-artifact__changed' : ''}`}>Page rotation: {String(content.rotation)}°</aside>
    </article>
  )
}

function ArtifactView({ format, content, highlighted }: { format: AgentDemoFormat; content: Record<string, unknown>; highlighted: boolean }) {
  if (format === 'xlsx') return <SheetArtifact content={content} highlighted={highlighted} />
  if (format === 'docx') return <DocxArtifact content={content} highlighted={highlighted} />
  if (format === 'pptx') return <PptxArtifact content={content} highlighted={highlighted} />
  return <PdfArtifact content={content} highlighted={highlighted} />
}

export default function AgentPage() {
  const [, setRouteTick] = useState(0)
  const tool = parseAgentTool()
  const format: AgentDemoFormat = agentFormatFromTool(tool)
  const [mode, setMode] = useState<AgentDemoMode>('safe')
  const [state, setState] = useState<WorkflowState>('ready')
  const [approved, setApproved] = useState(false)
  const [capabilities, setCapabilities] = useState<AgentCapability[]>([])
  const [inspection, setInspection] = useState<AgentInspection | null>(null)
  const [changeSet, setChangeSet] = useState<AgentDemoChangeSet | null>(null)
  const [preview, setPreview] = useState<AgentPreview | null>(null)
  const [diff, setDiff] = useState<AgentDiff | null>(null)
  const [validation, setValidation] = useState<AgentValidation | null>(null)
  const [receipt, setReceipt] = useState<AgentCommitReceipt | null>(null)
  const [verification, setVerification] = useState<AgentVerification | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [toolLog, setToolLog] = useState<AgentToolName[]>([])
  const scenario = useMemo(() => createAgentDemoScenario(format, mode), [format, mode])
  const toolMeta = AGENT_TOOLS.find((item) => item.tool === tool) ?? AGENT_TOOLS[0]
  const markTool = (name: AgentToolName) => setToolLog((log) => log.includes(name) ? log : [...log, name])

  const reset = () => {
    setState('ready')
    setApproved(false)
    setInspection(null)
    setChangeSet(null)
    setPreview(null)
    setDiff(null)
    setValidation(null)
    setReceipt(null)
    setVerification(null)
    setError(null)
    setToolLog([])
  }

  useEffect(() => {
    const syncTool = () => setRouteTick((tick) => tick + 1)
    window.addEventListener('hashchange', syncTool)
    return () => window.removeEventListener('hashchange', syncTool)
  }, [])

  useEffect(() => {
    reset()
    const { session } = createDemoSessionInput(format, mode)
    void session.capabilities().then((report) => setCapabilities(report.operations)).catch((reason: unknown) => {
      setError(reason instanceof Error ? reason.message : String(reason))
      setState('error')
    })
  }, [format, mode])

  const selectTool = (next: AgentTool) => {
    if (parseAgentTool() !== next) window.location.hash = agentHref(next)
  }

  const prepare = async () => {
    reset()
    setState('preparing')
    try {
      const { session, operations } = createDemoSessionInput(format, mode)
      const report = await session.capabilities()
      markTool('office.capabilities')
      setCapabilities(report.operations)
      const inspected = await session.inspect({ selection: format === 'xlsx' ? 'sheet-forecast!A1:D5' : 'document', maxItems: 20, maxBytes: 16_000 })
      markTool('office.inspect')
      setInspection(inspected)
      const planned = await session.plan(operations, { expectedRevision: inspected.revision })
      markTool('office.plan')
      setChangeSet(planned)
      const [nextPreview, nextDiff, nextValidation] = await Promise.all([planned.preview(), planned.diff(), planned.validate()])
      markTool('office.preview')
      markTool('office.diff')
      markTool('office.validate')
      setPreview(nextPreview)
      setDiff(nextDiff)
      setValidation(nextValidation)
      setState(nextValidation.ok ? 'awaiting-approval' : 'refused')
    } catch (reason: unknown) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setState('error')
    }
  }

  const commit = async () => {
    if (!changeSet || !inspection || !approved || !validation?.ok) return
    setState('committing')
    setError(null)
    try {
      const nextReceipt = await changeSet.commit({
        expectedRevision: inspection.revision,
        idempotencyKey: `demo-${changeSet.id}`,
        confirmation: 'approved',
      })
      markTool('office.commit')
      setReceipt(nextReceipt)
      const nextVerification = await changeSet.verify(nextReceipt)
      markTool('office.verify')
      setVerification(nextVerification)
      setState(nextVerification.ok ? 'verified' : 'error')
      if (!nextVerification.ok) setError('The committed output did not match its receipt.')
    } catch (reason: unknown) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setState('error')
    }
  }

  const shownArtifact = preview?.artifact ?? scenario.artifact
  const status = state === 'ready'
    ? 'Ready to inspect a bounded document selection.'
    : state === 'preparing'
      ? 'Inspecting and validating the proposed operations.'
      : state === 'awaiting-approval'
        ? 'Preflight passed. The source artifact is still unchanged.'
        : state === 'refused'
          ? 'Refused before write. No output was produced.'
          : state === 'committing'
            ? 'Applying the approved change set atomically.'
            : state === 'verified'
              ? `Verified ${receipt?.revision}. The output matches its receipt.`
              : 'The workflow stopped without changing the source.'

  return (
    <div className="ds">
    <section className="tool-page" data-demo-surface="agent" data-agent-tool={tool} aria-label={toolMeta.title}>
      <div className="agent-demo__toolbar workbench-toolbar ds-workstrip" role="toolbar" aria-label="Agent workflow controls">
        <DsSegment
          label="Office tool"
          value={tool}
          onChange={(id) => selectTool(id as AgentTool)}
          options={AGENT_TOOLS.map((item) => ({ id: item.tool, label: item.label }))}
        />
        <DsSegment
          label="Proposal type"
          value={mode}
          onChange={(id) => setMode(id as AgentDemoMode)}
          options={[
            { id: 'safe', label: 'Supported change' },
            { id: 'refusal', label: 'Refusal proof' },
          ]}
        />
        <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={state === 'preparing' || state === 'committing'} onClick={() => void prepare()}>Run agent</DsButton>
        <span className="agent-demo__status" data-state={state} role="status" aria-live="polite">{status}</span>
      </div>

      <div className="agent-request">
        <label htmlFor="agent-prompt">Agent request</label>
        <textarea id="agent-prompt" readOnly value={scenario.prompt} rows={2} />
        <small>Deterministic local proposal · lifecycle enforced by @injoffice/agent-tools · no model SDK or network request</small>
      </div>

      <ol className="agent-tool-log" aria-label="office.* tool calls">
        {AGENT_TOOL_METHODS.map((method) => (
          <li key={method} data-state={toolLog.includes(method) ? 'done' : 'waiting'}><code>{method}</code></li>
        ))}
      </ol>

      <ol className="agent-flight-recorder ds-timeline" aria-label="Agent change set stages">
        {STEPS.map((step) => {
          const current = stepState(step, state)
          return <li key={step} data-state={current}><i aria-hidden="true">{current === 'done' ? '✓' : current === 'refused' ? '!' : ''}</i><span>{step}</span></li>
        })}
      </ol>

      <div className="agent-demo__workspace ds-split">
        <section className="agent-demo__document ds-split-main" aria-labelledby="agent-artifact-title">
          <header>
            <div><span>Isolated {preview ? 'preview' : 'source'}</span><h2 id="agent-artifact-title">{scenario.artifact.name}</h2></div>
            <dl className="ds-proof">
              <div><dt>Artifact</dt><dd>{scenario.artifact.artifactId}</dd></div>
              <div><dt>Revision</dt><dd>{receipt?.revision ?? inspection?.revision ?? scenario.artifact.revision}</dd></div>
            </dl>
          </header>
          <div className={`agent-artifact agent-artifact--${format}`}>
            <ArtifactView format={format} content={shownArtifact.content} highlighted={Boolean(preview && validation?.ok)} />
          </div>
          <footer>
            <strong>Agent request</strong>
            <p>{scenario.prompt}</p>
            <small>Deterministic local proposal · lifecycle enforced by @injoffice/agent-tools · no model SDK or network request</small>
          </footer>
        </section>

        <aside className="agent-demo__inspector ds-split-side" aria-label="Agent evidence inspector">
          <section className="ds-panel">
            <header><div><span>Inspect</span><h2>Available operations</h2></div><b>{capabilities.length}</b></header>
            <div className="agent-capabilities">
              {capabilities.map((capability) => <span key={capability.operation} data-access={capability.destructive ? 'destructive' : capability.access}>{capability.operation}</span>)}
            </div>
          </section>

          <section className="ds-panel">
            <header><div><span>Plan</span><h2>{scenario.summary}</h2></div></header>
            {diff?.changes.length ? (
              <div className="agent-diff">
                {diff.changes.map((change) => <div className="ds-diff-row" key={change.target}><code>{change.target}</code><del>{change.before}</del><ins>{change.after}</ins></div>)}
              </div>
            ) : <p className="agent-demo__empty ds-muted">Prepare the change to see an isolated preview and semantic diff.</p>}
          </section>

          {validation && !validation.ok ? (
            <section className="agent-refusal ds-panel" role="alert">
              <header><div><span>Validate</span><h2>Change refused</h2></div></header>
              <DsCallout tone="refuse" title="Fail closed">
                Refused before write. No output was produced.
              </DsCallout>
              {validation.issues.map((issue) => <div key={issue.path}><strong>{issue.code}</strong><code>{issue.path}</code><p>{issue.message}</p></div>)}
            </section>
          ) : (
            <section className="agent-approval ds-panel">
              <header><div><span>Validate</span><h2>Review before commit</h2></div></header>
              {state === 'verified' ? <DsChip tone="green">Applied</DsChip> : null}
              <label className="ds-check">
                <input type="checkbox" checked={approved} disabled={state !== 'awaiting-approval'} onChange={(event) => setApproved(event.target.checked)} />
                I reviewed this exact diff and approve one atomic commit.
              </label>
              <DsButton variant={state === 'verified' ? 'green' : 'filled'} className="workbench-button workbench-button--primary" disabled={!approved || state !== 'awaiting-approval'} onClick={() => void commit()}>Commit approved change</DsButton>
            </section>
          )}

          <section className="agent-evidence ds-panel">
            <header><div><span>Commit</span><h2>{verification?.ok ? 'Output verified' : 'Execution record'}</h2></div>{verification?.ok && <DsChip tone="green">Pass</DsChip>}</header>
            <dl className="ds-proof">
              <div><dt>Source revision</dt><dd><code>{inspection?.revision ?? '—'}</code></dd></div>
              <div><dt>Source fingerprint</dt><dd><code>{inspection?.fingerprint ?? '—'}</code></dd></div>
              <div><dt>Output revision</dt><dd><code>{receipt?.revision ?? '—'}</code></dd></div>
              <div><dt>Output fingerprint</dt><dd><code>{verification?.fingerprint ?? '—'}</code></dd></div>
            </dl>
            {verification && <ul>{verification.evidence.map((item) => <li key={item}>{item}</li>)}</ul>}
          </section>
          {error && <p className="tool-error" role="alert">{error}</p>}
        </aside>
      </div>
    </section>
    </div>
  )
}
