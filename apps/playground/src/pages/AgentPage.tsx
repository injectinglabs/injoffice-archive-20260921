import { useCallback, useEffect, useId, useMemo, useRef, useState } from 'react'
import type { AgentToolCall, AgentToolCallResult, JsonObject } from '@injoffice/agent-tools'
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
import { createAgentDemoScenario, type AgentDemoFormat, type AgentDemoMode, type AgentDemoOperation } from '../agentDemoScenario'
import { AGENT_XLSX_NAME, createNativeAgentSessionInput } from '../agentXlsxDemo'
import { createNativeDocxAgentSessionInput } from '../agentDocxDemo'
import { createNativePptxAgentSessionInput } from '../agentPptxDemo'
import { createNativePdfAgentSessionInput } from '../agentPdfDemo'
import { requestMockAgentProposal } from '../mockAgentTransport'
import GuidedAgentTaskForm from '../components/GuidedAgentTaskForm'
import { createBrowserXlsxRoundTripRuntime } from '../xlsxRoundTripRuntime'
import { createBrowserDocxRoundTripRuntime } from '../docxRoundTripRuntime'
import { createBrowserPptxRoundTripRuntime } from '../pptxRoundTripRuntime'
import { AGENT_TOOLS, agentFormatFromTool, agentHref, parseAgentTool, parseSurface, type AgentTool } from '../route'
import { DsButton, DsCallout, DsChip, DsSegment } from '../design-system/primitives'
import '../design-system/live-tools.css'
import './AgentWorkflow.css'

type WorkflowState = 'ready' | 'preparing' | 'awaiting-approval' | 'refused' | 'committing' | 'verified' | 'unverified' | 'error'
type WorkflowStep = 'Inspect' | 'Plan' | 'Preview + diff' | 'Validate' | 'Approve' | 'Commit' | 'Verify'
type AgentToolName = 'office.capabilities' | 'office.inspect' | 'office.read' | 'office.plan' | 'office.preview' | 'office.diff' | 'office.validate' | 'office.commit' | 'office.verify'

const STEPS: WorkflowStep[] = ['Inspect', 'Plan', 'Preview + diff', 'Validate', 'Approve', 'Commit', 'Verify']
const AGENT_TOOL_METHODS: AgentToolName[] = ['office.capabilities', 'office.inspect', 'office.read', 'office.plan', 'office.preview', 'office.diff', 'office.validate', 'office.commit']
type TraceEntry = { request: AgentToolCall; response: AgentToolCallResult }
type SessionInput = ReturnType<typeof createDemoSessionInput> & {
  dispose?: () => void; download?: () => Blob | null
  approve?: (changeSetId: string) => Promise<void>
  propose?: (request: string) => Promise<AgentDemoOperation[]>
  trace?: () => TraceEntry[]
  simulateConcurrentEdit?: () => Promise<void>
  setVerificationFailure?: (enabled: boolean) => void
  stats?: () => { nativeWrites: number }
  proposalContext?: () => Promise<JsonObject>
  acceptProposal?: (proposal: unknown) => AgentDemoOperation[]
}

function SheetArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const headers = content.headers as string[]
  const rows = content.rows as Array<Array<string | number>>
  const changed = content.changedCell as { row: number; column: number } | undefined
  return (
    <div className="agent-artifact__sheet-scroll" role="region" aria-label="Workbook preview; scroll horizontally to see all columns" tabIndex={0}>
    <table className="agent-artifact__sheet ds-table">
      <thead><tr><th aria-label="Row number" />{headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
      <tbody>{rows.map((row, rowIndex) => <tr key={String(row[0])}><th>{rowIndex + 2}</th>{row.map((cell, cellIndex) => <td className={highlighted && rowIndex === (changed?.row ?? 3) && cellIndex === (changed?.column ?? 3) ? 'agent-artifact__changed' : undefined} key={`${rowIndex}-${cellIndex}`}>{cell}</td>)}</tr>)}</tbody>
    </table>
    </div>
  )
}

function DocxArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const blocks = content.blocks as Array<{ id: string; kind: string; text: string }>
  const titleBlock = blocks[0]?.text === content.title ? blocks[0] : undefined
  return (
    <article className="agent-artifact__doc ds-page">
      <h3 className={highlighted && titleBlock?.id === content.changedBlockId ? 'agent-artifact__changed' : undefined}>{String(content.title)}</h3>
      {blocks.filter((block) => block !== titleBlock).map((block) => <p className={highlighted && block.id === content.changedBlockId ? 'agent-artifact__changed' : undefined} key={block.id}>{block.text}</p>)}
    </article>
  )
}

function PptxArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const blocks = content.blocks as Array<{ id: string; text: string }>
  const titleBlock = blocks[0]?.text === content.title ? blocks[0] : undefined
  return (
    <article className="agent-artifact__slide ds-slide">
      <span>Extracted presentation text</span>
      <h3 className={highlighted && titleBlock?.id === content.changedBlockId ? 'agent-artifact__changed' : undefined}>{String(content.title)}</h3>
      <p>{String(content.subtitle)}</p>
      {blocks.filter((block) => block !== titleBlock).map((block) => <p key={block.id} className={highlighted && block.id === content.changedBlockId ? 'agent-artifact__changed' : undefined}>{block.text}</p>)}
    </article>
  )
}

function PdfArtifact({ content, highlighted }: { content: Record<string, unknown>; highlighted: boolean }) {
  const pages = content.pages as Array<{ page: number; rotation: number; width: number; height: number }>
  return (
    <article className="agent-artifact__pdf ds-pdf-sheet">
      <span>Parsed PDF page metadata · {String(content.pageCount)} pages</span>
      <h3>{String(content.title)}</h3>
      <div className="agent-pdf-page-scroll" role="region" aria-label="PDF page metadata" tabIndex={0}>
        <table className="ds-table"><thead><tr><th>Page</th><th>Rotation</th><th>Size (pt)</th></tr></thead>
          <tbody>{pages.map((page) => <tr key={page.page} className={highlighted && page.page === content.selectedPage ? 'agent-artifact__changed' : undefined}><th>{page.page}</th><td>{page.rotation}°</td><td>{Math.round(page.width)} × {Math.round(page.height)}</td></tr>)}</tbody>
        </table>
      </div>
      <p>This is a page-metadata preview, not a rendered PDF page. The verified download contains the real document.</p>
    </article>
  )
}

function ArtifactView({ format, content, highlighted }: { format: AgentDemoFormat; content: Record<string, unknown>; highlighted: boolean }) {
  if (format === 'xlsx') return <SheetArtifact content={content} highlighted={highlighted} />
  if (format === 'docx') return <DocxArtifact content={content} highlighted={highlighted} />
  if (format === 'pptx') return <PptxArtifact content={content} highlighted={highlighted} />
  return <PdfArtifact content={content} highlighted={highlighted} />
}

export default function AgentPage({ fixedTool }: { fixedTool?: AgentTool } = {}) {
  const [routedTool, setRoutedTool] = useState<AgentTool>(() => parseAgentTool())
  useEffect(() => {
    if (fixedTool !== undefined) return
    const syncTool = () => {
      if (parseSurface() === 'agent') setRoutedTool(parseAgentTool())
    }
    syncTool()
    window.addEventListener('hashchange', syncTool)
    return () => window.removeEventListener('hashchange', syncTool)
  }, [fixedTool])
  const tool = fixedTool ?? routedTool
  return <AgentWorkflow key={tool} tool={tool} fixedTool={fixedTool !== undefined} />
}

function AgentWorkflow({ tool, fixedTool }: { tool: AgentTool; fixedTool: boolean }) {
  const instanceId = useId()
  const promptId = `agent-prompt-${instanceId}`
  const artifactTitleId = `agent-artifact-title-${instanceId}`
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
  const [sessionInput, setSessionInput] = useState<SessionInput | null>(null)
  const [reload, setReload] = useState(0)
  const [downloadURL, setDownloadURL] = useState('')
  const [prompt, setPrompt] = useState('Mark Security as Ready')
  const [trace, setTrace] = useState<TraceEntry[]>([])
  const [nativeWrites, setNativeWrites] = useState(0)
  const [sourceChanged, setSourceChanged] = useState(false)
  const [failVerification, setFailVerification] = useState(false)
  const [safetyBusy, setSafetyBusy] = useState(false)
  const [retryResult, setRetryResult] = useState('')
  const [proposalExchange, setProposalExchange] = useState<{ request: JsonObject; response: unknown } | null>(null)
  const [proposalContext, setProposalContext] = useState<JsonObject | null>(null)
  const [advancedRequest, setAdvancedRequest] = useState(false)
  const proposalAbort = useRef<AbortController | null>(null)
  const generation = useRef(0)
  const fallbackScenario = useMemo(() => {
    const sample = createAgentDemoScenario(format, mode)
    if (format !== 'xlsx') return sample
    return { ...sample, artifact: { ...sample.artifact, name: AGENT_XLSX_NAME, revision: 'Not loaded' },
      prompt: mode === 'safe' ? 'Mark Security as Ready' : 'Execute an unsupported workbook macro.',
      summary: 'Load the real workbook to inspect the exact source and proposed change.',
    }
  }, [format, mode])
  const scenario = sessionInput?.scenario ?? fallbackScenario
  const toolMeta = AGENT_TOOLS.find((item) => item.tool === tool) ?? AGENT_TOOLS[0]
  const markTool = (name: AgentToolName) => setToolLog((log) => log.includes(name) ? log : [...log, name])
  const refreshTrace = (input: SessionInput | null = sessionInput) => {
    setTrace(input?.trace?.() ?? [])
    setNativeWrites(input?.stats?.().nativeWrites ?? 0)
  }

  const reset = useCallback(() => {
    proposalAbort.current?.abort()
    generation.current += 1
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
    setRetryResult('')
    setProposalExchange(null)
  }, [])

  useEffect(() => {
    const blob = verification?.ok ? sessionInput?.download?.() : null
    if (!blob) { setDownloadURL(''); return }
    const url = URL.createObjectURL(blob)
    setDownloadURL(url)
    return () => URL.revokeObjectURL(url)
  }, [verification, sessionInput])

  useEffect(() => {
    reset()
    setSessionInput(null)
    setCapabilities([])
    setSourceChanged(false)
    setFailVerification(false)
    setSafetyBusy(false)
    setTrace([])
    setNativeWrites(0)
    setProposalContext(null)
    setAdvancedRequest(false)
    let cancelled = false
    let loaded: SessionInput | undefined
    const xlsxRuntime = format === 'xlsx' ? createBrowserXlsxRoundTripRuntime() : undefined
    const docxRuntime = format === 'docx' ? createBrowserDocxRoundTripRuntime() : undefined
    const pptxRuntime = format === 'pptx' ? createBrowserPptxRoundTripRuntime() : undefined
    const terminateLoading = () => { xlsxRuntime?.terminate(); docxRuntime?.terminate(); pptxRuntime?.terminate() }
    const initialize = async () => {
      try {
        loaded = format === 'xlsx' ? await createNativeAgentSessionInput(mode, xlsxRuntime)
          : format === 'docx' ? await createNativeDocxAgentSessionInput(mode, docxRuntime)
          : format === 'pptx' ? await createNativePptxAgentSessionInput(mode, pptxRuntime)
          : await createNativePdfAgentSessionInput(mode)
        if (cancelled) { loaded.dispose?.(); return }
        const report = await loaded.session.capabilities()
        if (cancelled) return
        setSessionInput(loaded)
        setCapabilities(report.operations)
        setPrompt(loaded.scenario.prompt)
        refreshTrace(loaded)
      } catch (reason: unknown) {
        loaded?.dispose?.()
        terminateLoading()
        if (cancelled) return
        setError(reason instanceof Error ? reason.message : String(reason))
        setState('error')
      }
    }
    void initialize()
    return () => { cancelled = true; generation.current += 1; proposalAbort.current?.abort(); loaded?.dispose?.(); terminateLoading() }
  }, [format, mode, reload])

  useEffect(() => {
    setProposalContext(null)
    if (mode !== 'safe' || !sessionInput?.proposalContext) return
    let cancelled = false
    const connect = async () => {
      try {
        const context = await sessionInput.proposalContext!()
        if (cancelled) return
        setProposalContext(context)
        refreshTrace()
      } catch (reason) {
        if (!cancelled) {
          const message = reason instanceof Error ? reason.message : String(reason)
          setError(message)
          setState('error')
        }
      }
    }
    void connect()
    return () => { cancelled = true }
  }, [mode, sessionInput])

  const selectTool = (next: AgentTool) => {
    if (parseAgentTool() !== next) window.location.hash = agentHref(next)
  }

  const reloadSample = () => { reset(); setSessionInput(null); setReload((value) => value + 1) }
  const selectMode = (next: AgentDemoMode) => {
    if (next === mode) return
    reset()
    setSessionInput(null)
    setMode(next)
  }

  const prepare = async () => {
    if (!sessionInput) return
    reset()
    const run = generation.current
    setState('preparing')
    try {
      const { session } = sessionInput
      let operations = sessionInput.operations
      if (mode === 'safe' && sessionInput.propose) {
        if (!proposalContext || !sessionInput.acceptProposal) throw new Error('The mock context is not ready. Reload the sample and try again.')
        const controller = new AbortController()
        proposalAbort.current = controller
        const { capabilities: sharedCapabilities, ...context } = proposalContext
        const request = { request: prompt, context, capabilities: sharedCapabilities } as JsonObject
        const proposed = await requestMockAgentProposal(request, controller.signal)
        if (run !== generation.current) return
        setProposalExchange({ request, response: proposed })
        operations = sessionInput.acceptProposal(proposed)
      }
      if (run !== generation.current) return
      refreshTrace()
      const report = await session.capabilities()
      if (run !== generation.current) return
      markTool('office.capabilities')
      setCapabilities(report.operations)
      const inspected = await session.inspect({ selection: format === 'xlsx' ? 'A1:F8' : 'document', maxItems: 20, maxBytes: 16_000 })
      if (run !== generation.current) return
      markTool('office.inspect')
      setInspection(inspected)
      const planned = await session.plan(operations, { expectedRevision: inspected.revision })
      if (run !== generation.current) return
      markTool('office.plan')
      setChangeSet(planned)
      const [nextPreview, nextDiff, nextValidation] = await Promise.all([planned.preview(), planned.diff(), planned.validate()])
      if (run !== generation.current) return
      markTool('office.preview')
      markTool('office.diff')
      markTool('office.validate')
      setPreview(nextPreview)
      setDiff(nextDiff)
      setValidation(nextValidation)
      setState(nextValidation.ok ? 'awaiting-approval' : 'refused')
    } catch (reason: unknown) {
      if (run !== generation.current) return
      setError(reason instanceof Error ? reason.message : String(reason))
      setState('error')
    } finally {
      if (run === generation.current) refreshTrace()
    }
  }

  const commit = async () => {
    if (!changeSet || !inspection || !approved || !validation?.ok) return
    setState('committing')
    setError(null)
    const run = generation.current
    try {
      // This function is reached from a trusted UI action, never a tool call.
      await sessionInput?.approve?.(changeSet.id)
      if (run !== generation.current) return
      sessionInput?.setVerificationFailure?.(failVerification)
      const nextReceipt = await changeSet.commit({
        expectedRevision: inspection.revision,
        idempotencyKey: `demo-${changeSet.id}`,
        confirmation: 'approved',
      })
      if (run !== generation.current) return
      markTool('office.commit')
      setReceipt(nextReceipt)
      const nextVerification = await changeSet.verify(nextReceipt)
      if (run !== generation.current) return
      setVerification(nextVerification)
      setState(nextVerification.ok ? 'verified' : 'unverified')
      if (!nextVerification.ok) setError('Write completed; verification failed. The commit receipt exists, but no verified download is available. Reload the sample before trying another edit.')
    } catch (reason: unknown) {
      if (run !== generation.current) return
      setError(reason instanceof Error ? reason.message : String(reason))
      setState('error')
    } finally {
      if (run === generation.current) refreshTrace()
    }
  }

  const simulateConcurrentEdit = async () => {
    if (!sessionInput?.simulateConcurrentEdit || sourceChanged || safetyBusy) return
    const run = generation.current
    setSafetyBusy(true)
    try {
      await sessionInput.simulateConcurrentEdit()
      if (run === generation.current) setSourceChanged(true)
    } catch (reason) {
      if (run === generation.current) setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      if (run === generation.current) { setSafetyBusy(false); refreshTrace() }
    }
  }

  const retryCommit = async () => {
    if (!changeSet || !inspection || !receipt || !sessionInput?.stats || safetyBusy) return
    const run = generation.current
    const before = sessionInput.stats().nativeWrites
    setSafetyBusy(true)
    try {
      const repeated = await changeSet.commit({ expectedRevision: inspection.revision, idempotencyKey: `demo-${changeSet.id}`, confirmation: 'approved' })
      if (run !== generation.current) return
      const after = sessionInput.stats().nativeWrites
      setRetryResult(repeated.fingerprint === receipt.fingerprint && before === after ? 'Same receipt returned. No additional native write.' : 'Retry returned a different result; inspect the tool trace.')
    } catch (reason) {
      if (run === generation.current) setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      if (run === generation.current) { setSafetyBusy(false); refreshTrace() }
    }
  }

  const shownArtifact = preview?.artifact ?? scenario.artifact
  const shownContent = verification?.ok && receipt?.content ? receipt.content : shownArtifact.content
  const status = !sessionInput && state !== 'error' ? 'Loading the bundled sample and browser engine…' : state === 'ready'
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
              : state === 'unverified' ? 'Write completed; verification failed. No verified download is available.'
              : 'The workflow stopped. See the error details below; reload the sample to try again.'
  const fileType = format.toUpperCase()
  const boundary = `Real ${fileType} preview, approval, file write, and verification. No language model, provider credentials, model service, or external AI service.`
  const requestHelp = format === 'xlsx'
    ? 'Try “Mark Mobile as On track” or “Mark Security as Ready”. Supports one workstream status edit at a time.'
    : format === 'pdf' ? 'Try “Rotate page 2 by 90 degrees” or “Rotate page 1 by 180 degrees”. Supports one page rotation at a time.'
    : 'Use Replace "exact current text" with "replacement text". Supports one uniquely identified editable text target at a time; copy the current text from the preview.'
  const taskDisabled = state === 'preparing' || state === 'committing' || state === 'verified' || state === 'unverified' || sourceChanged || safetyBusy
  const updateRequest = useCallback((request: string) => { reset(); setPrompt(request) }, [reset])
  const stageStates: Record<WorkflowStep, 'done' | 'active' | 'waiting' | 'refused'> = {
    Inspect: inspection ? 'done' : state === 'preparing' ? 'active' : 'waiting',
    Plan: changeSet ? 'done' : inspection && state === 'preparing' ? 'active' : 'waiting',
    'Preview + diff': preview && diff ? 'done' : changeSet && state === 'preparing' ? 'active' : 'waiting',
    Validate: validation ? validation.ok ? 'done' : 'refused' : changeSet && state === 'preparing' ? 'active' : 'waiting',
    Approve: approved ? 'done' : state === 'awaiting-approval' ? 'active' : 'waiting',
    Commit: receipt ? 'done' : state === 'committing' ? 'active' : 'waiting',
    Verify: verification ? verification.ok ? 'done' : 'refused' : receipt && state === 'committing' ? 'active' : 'waiting',
  }
  const taskSteps = [
    { label: 'Choose a task', detail: 'Set a small, specific change.', state: state === 'ready' || state === 'error' ? 'active' : 'done' },
    { label: 'Inspect and preview', detail: 'Read the file and compare the change.', state: validation ? validation.ok ? 'done' : 'refused' : state === 'preparing' ? 'active' : 'waiting' },
    { label: 'Review and approve', detail: 'Only you can allow the file write.', state: receipt ? 'done' : state === 'awaiting-approval' || state === 'committing' ? 'active' : 'waiting' },
    { label: 'Verify and download', detail: 'Reopen the edited file to check it.', state: verification ? verification.ok ? 'done' : 'refused' : state === 'committing' ? 'active' : 'waiting' },
  ]

  return (
    <div className="ds">
    <section className="tool-page" data-demo-surface="agent" data-agent-tool={tool} aria-label={toolMeta.title}>
      <header className="agent-task-intro" data-agent-mock>
        <h2>Simulated agent · real document operations</h2>
        <p data-agent-boundary>{boundary}</p>
        <p>Choose a task for the sample file. Review what will change, then approve the edit and download the verified result.</p>
      </header>
      <ol className="agent-task-progress" data-agent-progress aria-label="Document task progress">
        {taskSteps.map((step, index) => <li key={step.label} data-state={step.state} aria-current={step.state === 'active' ? 'step' : undefined}><span aria-hidden="true">{step.state === 'done' ? '✓' : index + 1}</span><div><strong>{step.label}</strong><small>{step.detail}</small></div></li>)}
      </ol>
      <section className="agent-task-editor" aria-label="Choose a document task">
        {!fixedTool && <DsSegment
          label="Office tool"
          value={tool}
          onChange={(id) => selectTool(id as AgentTool)}
          options={AGENT_TOOLS.map((item) => ({ id: item.tool, label: item.label }))}
        />}
        {mode === 'safe' && advancedRequest ? <div className="agent-advanced-notice" role="status"><strong>Advanced request active</strong><p>{prompt || 'Enter a bounded request in Technical details below.'}</p><small>Task fields are paused. This exact request will be proposed by the deterministic mock.</small></div> : mode === 'safe' && proposalContext && sessionInput ? <GuidedAgentTaskForm key={`${reload}-${mode}`} tool={tool} context={proposalContext} defaultRequest={sessionInput.scenario.prompt} disabled={taskDisabled} onRequestChange={updateRequest} /> : <p>{mode === 'refusal' ? 'Safety test: propose an unsupported operation and confirm that it is refused before any write.' : 'Reading the sample to find the available task targets…'}</p>}
        <div className="agent-demo__toolbar workbench-toolbar ds-workstrip" role="toolbar" aria-label="Agent workflow controls">
        <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={(!sessionInput && state !== 'error') || state === 'preparing' || state === 'committing' || safetyBusy || (sourceChanged && state === 'awaiting-approval') || (mode === 'safe' && state !== 'verified' && state !== 'unverified' && state !== 'error' && (!proposalContext || !prompt.trim()))} onClick={() => state === 'verified' || state === 'unverified' || state === 'error' ? reloadSample() : void prepare()}>{state === 'verified' || state === 'unverified' || state === 'error' ? 'Reload sample' : 'Run agent'}</DsButton>
        <span className="agent-demo__status" data-state={state} role="status" aria-live="polite">{status}</span>
        </div>
      </section>

      <div className="agent-demo__workspace ds-split">
        <section className="agent-demo__document ds-split-main" aria-labelledby={artifactTitleId}>
          <header>
            <div><span>{verification?.ok ? 'Reopened output projection' : `Isolated ${preview ? 'preview' : 'source'}`}</span><h2 id={artifactTitleId}>{scenario.artifact.name}</h2></div>
          </header>
          <div className={`agent-artifact agent-artifact--${format}`}>
            {sessionInput ? <ArtifactView format={format} content={shownContent} highlighted={Boolean(preview && validation?.ok)} /> : <p role="status">{state === 'error' ? 'Sample unavailable. Reload to retry.' : 'Opening the sample…'}</p>}
          </div>
          <footer>
            <strong>Agent request</strong>
            <p>{mode === 'safe' ? prompt : scenario.prompt}</p>
            <small>{boundary}</small>
            <p className="agent-approval-note">{format === 'pdf' ? 'PDF page metadata comes from parsing the actual file bytes.' : format === 'xlsx' ? 'The table is a bounded workbook projection, not a rendered Excel page.' : 'This is extracted document text, not a full-fidelity page or slide rendering.'}</p>
          </footer>
        </section>

        <aside className="agent-demo__inspector ds-split-side" aria-label="Agent evidence inspector">
          <section className="ds-panel">
            <header><div><h2>{mode === 'safe' ? `Proposed ${fileType} change` : scenario.summary}</h2></div></header>
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
              {validation.issues.map((issue, index) => <div key={`${issue.code}:${issue.path}:${index}`}><strong>{issue.code}</strong><code>{issue.path}</code><p>{issue.message}</p></div>)}
            </section>
          ) : (
            <section className="agent-approval ds-panel">
              <header><div><h2>Review and approve</h2></div></header>
              {state === 'verified' ? <DsChip tone="green">Applied</DsChip> : null}
              <label className="ds-check">
                <input type="checkbox" checked={approved} disabled={state !== 'awaiting-approval'} onChange={(event) => setApproved(event.target.checked)} />
                I reviewed this exact diff and approve one atomic commit.
              </label>
              <DsButton variant={state === 'verified' ? 'green' : 'filled'} className="workbench-button workbench-button--primary" disabled={!approved || state !== 'awaiting-approval' || safetyBusy} onClick={() => void commit()}>Commit approved change</DsButton>
              <p className="agent-approval-note">The original file is unchanged until you approve. Editing the task clears this approval.</p>
            </section>
          )}

          <section className="agent-evidence ds-panel">
            <header><div><h2>{verification?.ok ? `${fileType} output verified` : 'Verify and download'}</h2></div>{verification?.ok && <DsChip tone="green">Pass</DsChip>}</header>
            <p>{verification?.ok ? 'The edited file was reopened and checked. Your verified download is ready.' : state === 'unverified' ? 'The write completed, but its result could not be verified. Download is blocked; reload the sample to start again.' : 'After approval, the real file is edited and reopened for verification. Only a verified result can be downloaded.'}</p>
            {state === 'verified' && verification?.ok && downloadURL && <a data-agent-download className="workbench-button ds-btn ds-btn--filled" href={downloadURL} download={`${scenario.artifact.name.replace(/\.[^.]+$/, '')}-approved.${format}`}>Download verified .{format}</a>}
          </section>
          {sourceChanged && <p role="status">The source has changed. Try committing the reviewed plan: its stale revision must be rejected.</p>}
          {failVerification && <p role="status">Safety test enabled: verification will fail after the write, so no download will be offered.</p>}
          {error && <p className="tool-error" role="alert">{error}</p>}
        </aside>
      </div>
      <details data-agent-technical className="agent-technical">
        <summary>Technical details</summary>
        <div className="agent-request">
          <label className="ds-check"><input type="checkbox" data-agent-advanced-request checked={advancedRequest} disabled={taskDisabled || mode !== 'safe'} onChange={(event) => { reset(); setAdvancedRequest(event.target.checked) }} />Use an advanced request instead of task fields</label>
          <label htmlFor={promptId}>Bounded mock request</label>
          <textarea id={promptId} data-agent-request maxLength={2000} readOnly={!advancedRequest || mode !== 'safe'} disabled={taskDisabled} value={mode === 'safe' ? prompt : scenario.prompt} onChange={(event) => updateRequest(event.target.value)} rows={2} />
          <p className="agent-request-help">{requestHelp} This is a deterministic mock, not an open-ended chat.</p>
          <p className="agent-request-help">The built-in mock returns a proposal from the inspected document. {import.meta.env.DEV ? 'It runs on this local demo server.' : 'On this static site, the response is simulated in your browser.'} No provider is contacted. Preview, approval, file editing, and verification are real.</p>
          {proposalExchange && <details data-agent-proposal-trace><summary>Mock proposal request and response</summary><pre className="ds-code" tabIndex={0}>{JSON.stringify(proposalExchange, null, 2)}</pre></details>}
        </div>
        <ol className="agent-tool-log" aria-label="Actual office.* tool calls">
          {AGENT_TOOL_METHODS.map((method) => <li key={method} data-state={(sessionInput?.trace ? trace.some((entry) => entry.request.method === method && entry.response.ok) : toolLog.includes(method)) ? 'done' : 'waiting'}><code>{method}</code></li>)}
        </ol>
        <ol className="agent-flight-recorder ds-timeline" aria-label="Agent change set stages">
          {STEPS.map((step) => <li key={step} data-state={stageStates[step]}><i aria-hidden="true">{stageStates[step] === 'done' ? '✓' : stageStates[step] === 'refused' ? '!' : ''}</i><span>{step}</span></li>)}
        </ol>
        <section className="agent-technical-evidence">
          <h3>Capabilities and verification evidence</h3>
          <div className="agent-capabilities">{capabilities.map((capability) => <span key={capability.operation} data-access={capability.destructive ? 'destructive' : capability.access}>{capability.operation}</span>)}</div>
          <dl className="ds-proof">
            <div><dt>Artifact</dt><dd>{scenario.artifact.artifactId}</dd></div>
            <div><dt>Source revision</dt><dd><code>{inspection?.revision ?? '—'}</code></dd></div>
            <div><dt>Source fingerprint</dt><dd><code>{inspection?.fingerprint ?? '—'}</code></dd></div>
            <div><dt>Output revision</dt><dd><code>{receipt?.revision ?? '—'}</code></dd></div>
            <div><dt>Output fingerprint</dt><dd><code>{verification?.fingerprint ?? '—'}</code></dd></div>
          </dl>
          {verification && <ul>{verification.evidence.map((item) => <li key={item}>{item}</li>)}</ul>}
          <p>Verification comes from the committed receipt, not a second <code>office.verify</code> call. Approval is stored by the host for this exact change set. A tool argument saying “approved” cannot grant permission.</p>
        </section>
        <details data-agent-safety-details className="agent-safety" aria-label="Safety scenarios">
          <summary>Try the safety boundaries</summary>
          <DsSegment label="Proposal type" value={mode} onChange={(id) => { if (state !== 'preparing' && state !== 'committing' && !safetyBusy) selectMode(id as AgentDemoMode) }} options={[{ id: 'safe', label: 'Supported change' }, { id: 'refusal', label: 'Refusal proof' }]} />
          <p>Source file writes: <span data-agent-native-writes>{nativeWrites}</span>. Isolated preview writes are not counted.</p>
          {sessionInput?.simulateConcurrentEdit && <DsButton data-agent-concurrent-edit variant="outlined" disabled={state !== 'awaiting-approval' || sourceChanged || safetyBusy} onClick={() => void simulateConcurrentEdit()}>Simulate another editor changing the source</DsButton>}
          {sessionInput?.setVerificationFailure && <><label className="ds-check"><input data-agent-fail-verification type="checkbox" checked={failVerification} disabled={state !== 'awaiting-approval' || safetyBusy} onChange={(event) => setFailVerification(event.target.checked)} />Inject a verification-read failure after the write</label><p>This explicitly simulates a readback failure; it does not corrupt the sample file.</p></>}
          <DsButton data-agent-retry variant="outlined" disabled={state !== 'verified' || safetyBusy} onClick={() => void retryCommit()}>Retry the same commit</DsButton>
          {retryResult && <p role="status">{retryResult}</p>}
        </details>
      <details data-agent-trace className="agent-dispatch-trace">
        <summary>Actual tool requests and results ({trace.length})</summary>
        <p>Requests pass through <code>createAgentToolDispatcher</code> and the public {fileType} adapter. The final commit result contains post-write verification. Read and preview results are bounded document projections.</p>
        {trace.map((entry) => <details key={entry.request.requestId}>
          <summary>{entry.request.method} · {entry.response.ok ? 'Completed' : 'Refused or failed'}</summary>
          <pre className="ds-code" tabIndex={0}>{JSON.stringify(entry, null, 2)}</pre>
        </details>)}
      </details>
      </details>
    </section>
    </div>
  )
}
