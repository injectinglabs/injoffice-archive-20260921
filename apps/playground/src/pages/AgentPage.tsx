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

type WorkflowState = 'ready' | 'preparing' | 'awaiting-approval' | 'refused' | 'committing' | 'verified' | 'error'
type WorkflowStep = 'Inspect' | 'Plan' | 'Preview + diff' | 'Validate' | 'Approve' | 'Commit' | 'Verify'

const STEPS: WorkflowStep[] = ['Inspect', 'Plan', 'Preview + diff', 'Validate', 'Approve', 'Commit', 'Verify']

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
    <table className="agent-artifact__sheet">
      <thead><tr><th aria-label="Row number" />{headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
      <tbody>{rows.map((row, rowIndex) => <tr key={String(row[0])}><th>{rowIndex + 2}</th>{row.map((cell, cellIndex) => <td className={highlighted && rowIndex === 3 && cellIndex === 3 ? 'agent-artifact__changed' : undefined} key={`${rowIndex}-${cellIndex}`}>{cell}</td>)}</tr>)}</tbody>
    </table>
  )
}

function DocxArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const blocks = content.blocks as Array<{ id: string; kind: string; text: string }>
  return (
    <article className="agent-artifact__doc">
      <h3>{String(content.title)}</h3>
      {blocks.map((block) => <p className={highlighted && block.id === 'block-summary' ? 'agent-artifact__changed' : undefined} key={block.id}>{block.text}</p>)}
    </article>
  )
}

function PptxArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const metric = content.metric as { label: string; value: string }
  return (
    <article className="agent-artifact__slide">
      <span>Board update</span>
      <h3>{String(content.title)}</h3>
      <p>{String(content.subtitle)}</p>
      <div className={`agent-artifact__metric${highlighted ? ' agent-artifact__changed' : ''}`}><strong>{metric.value}</strong><small>{metric.label}</small></div>
    </article>
  )
}

function PdfArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  return (
    <article className="agent-artifact__pdf">
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
  const [format, setFormat] = useState<AgentDemoFormat>('xlsx')
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
  const scenario = useMemo(() => createAgentDemoScenario(format, mode), [format, mode])

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
  }

  useEffect(() => {
    reset()
    const { session } = createDemoSessionInput(format, mode)
    void session.capabilities().then((report) => setCapabilities(report.operations)).catch((reason: unknown) => {
      setError(reason instanceof Error ? reason.message : String(reason))
      setState('error')
    })
  }, [format, mode])

  const prepare = async () => {
    reset()
    setState('preparing')
    try {
      const { session, operations } = createDemoSessionInput(format, mode)
      const report = await session.capabilities()
      setCapabilities(report.operations)
      const inspected = await session.inspect({ selection: format === 'xlsx' ? 'sheet-forecast!A1:D5' : 'document', maxItems: 20, maxBytes: 16_000 })
      setInspection(inspected)
      const planned = await session.plan(operations, { expectedRevision: inspected.revision })
      setChangeSet(planned)
      const [nextPreview, nextDiff, nextValidation] = await Promise.all([planned.preview(), planned.diff(), planned.validate()])
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
      setReceipt(nextReceipt)
      const nextVerification = await changeSet.verify(nextReceipt)
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
    <section className="tool-page" data-demo-surface="agent" aria-label="Agent change set workbench">
      <div className="agent-demo__toolbar workbench-toolbar" role="toolbar" aria-label="Agent workflow controls">
        <label className="tool-field">Document
          <select value={format} onChange={(event) => setFormat(event.target.value as AgentDemoFormat)}>
            <option value="xlsx">Spreadsheet · XLSX</option>
            <option value="docx">Document · DOCX</option>
            <option value="pptx">Presentation · PPTX</option>
            <option value="pdf">Portable document · PDF</option>
          </select>
        </label>
        <span className="tool-segment" role="group" aria-label="Proposal type">
          <button type="button" aria-pressed={mode === 'safe'} onClick={() => setMode('safe')}>Supported change</button>
          <button type="button" aria-pressed={mode === 'refusal'} onClick={() => setMode('refusal')}>Refusal proof</button>
        </span>
        <button className="workbench-button workbench-button--primary" type="button" disabled={state === 'preparing' || state === 'committing'} onClick={() => void prepare()}>Prepare change</button>
        <span className="agent-demo__status" data-state={state} role="status" aria-live="polite">{status}</span>
      </div>

      <ol className="agent-flight-recorder" aria-label="Agent change set stages">
        {STEPS.map((step) => {
          const current = stepState(step, state)
          return <li key={step} data-state={current}><i aria-hidden="true">{current === 'done' ? '✓' : current === 'refused' ? '!' : ''}</i><span>{step}</span></li>
        })}
      </ol>

      <div className="agent-demo__workspace">
        <section className="agent-demo__document" aria-labelledby="agent-artifact-title">
          <header>
            <div><span>Isolated {preview ? 'preview' : 'source'}</span><h2 id="agent-artifact-title">{scenario.artifact.name}</h2></div>
            <dl>
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

        <aside className="agent-demo__inspector" aria-label="Agent evidence inspector">
          <section>
            <header><div><span>Discover</span><h2>Available operations</h2></div><b>{capabilities.length}</b></header>
            <div className="agent-capabilities">
              {capabilities.map((capability) => <span key={capability.operation} data-access={capability.destructive ? 'destructive' : capability.access}>{capability.operation}</span>)}
            </div>
          </section>

          <section>
            <header><div><span>Proposed change</span><h2>{scenario.summary}</h2></div></header>
            {diff?.changes.length ? (
              <div className="agent-diff">
                {diff.changes.map((change) => <div key={change.target}><code>{change.target}</code><del>{change.before}</del><ins>{change.after}</ins></div>)}
              </div>
            ) : <p className="agent-demo__empty">Prepare the change to see an isolated preview and semantic diff.</p>}
          </section>

          {validation && !validation.ok ? (
            <section className="agent-refusal" role="alert">
              <header><div><span>Fail closed</span><h2>Change refused</h2></div></header>
              {validation.issues.map((issue) => <div key={issue.path}><strong>{issue.code}</strong><code>{issue.path}</code><p>{issue.message}</p></div>)}
            </section>
          ) : (
            <section className="agent-approval">
              <header><div><span>Human boundary</span><h2>Review before commit</h2></div></header>
              <label>
                <input type="checkbox" checked={approved} disabled={state !== 'awaiting-approval'} onChange={(event) => setApproved(event.target.checked)} />
                I reviewed this exact diff and approve one atomic commit.
              </label>
              <button className="workbench-button workbench-button--primary" type="button" disabled={!approved || state !== 'awaiting-approval'} onClick={() => void commit()}>Commit approved change</button>
            </section>
          )}

          <section className="agent-evidence">
            <header><div><span>Machine evidence</span><h2>{verification?.ok ? 'Output verified' : 'Execution record'}</h2></div>{verification?.ok && <b>Pass</b>}</header>
            <dl>
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
  )
}
