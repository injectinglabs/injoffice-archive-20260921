import { useEffect, useMemo, useRef, useState } from 'react'
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
import { createBrowserXlsxRoundTripRuntime } from '../xlsxRoundTripRuntime'
import { createBrowserDocxRoundTripRuntime } from '../docxRoundTripRuntime'
import { createBrowserPptxRoundTripRuntime } from '../pptxRoundTripRuntime'
import { AGENT_TOOLS, agentFormatFromTool, agentHref, parseAgentTool, type AgentTool } from '../route'
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

function stepState(step: WorkflowStep, state: WorkflowState): 'done' | 'active' | 'waiting' | 'refused' {
  const progress: Record<WorkflowState, number> = {
    ready: 0,
    preparing: 2,
    'awaiting-approval': 4,
    refused: 3,
    committing: 5,
    verified: 7,
    unverified: 6,
    error: 0,
  }
  const index = STEPS.indexOf(step)
  if (state === 'refused' && step === 'Validate') return 'refused'
  if (state === 'unverified' && step === 'Verify') return 'refused'
  if (index < progress[state]) return 'done'
  if (index === progress[state]) return 'active'
  return 'waiting'
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

export default function AgentPage() {
  const [, setRouteTick] = useState(0)
  const tool = parseAgentTool()
  useEffect(() => {
    const syncTool = () => setRouteTick((tick) => tick + 1)
    window.addEventListener('hashchange', syncTool)
    return () => window.removeEventListener('hashchange', syncTool)
  }, [])
  return <AgentWorkflow key={tool} tool={tool} />
}

function AgentWorkflow({ tool }: { tool: AgentTool }) {
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
  const [proposalSource, setProposalSource] = useState<'mock' | 'local' | 'live'>('mock')
  const [proposalExchange, setProposalExchange] = useState<{ request: JsonObject; response: unknown } | null>(null)
  const [liveContext, setLiveContext] = useState<JsonObject | null>(null)
  const [liveHost, setLiveHost] = useState('')
  const [liveNotice, setLiveNotice] = useState('')
  const [consent, setConsent] = useState(false)
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

  const reset = () => {
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
  }

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
    setConsent(false)
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
    setLiveContext(null)
    setLiveHost('')
    setConsent(false)
    if (mode !== 'safe' || proposalSource === 'local' || !sessionInput?.proposalContext) return
    const controller = new AbortController()
    let cancelled = false
    setLiveNotice(proposalSource === 'mock' ? 'Reading the sample for the built-in mock…' : 'Checking the local proposal host…')
    const connect = async () => {
      try {
        if (proposalSource === 'mock') {
          const context = await sessionInput.proposalContext!()
          if (cancelled) return
          setLiveContext(context)
          setLiveNotice('Mock ready. No language model, API key, or external service is used.')
          refreshTrace()
          return
        }
        const response = await fetch('/api/agent/proposal-status', { signal: controller.signal })
        if (!response.ok || !response.headers.get('content-type')?.includes('application/json')) throw new Error('Live proposals require a configured local host. This deployment supports local mode.')
        const status = await response.json() as { configured?: boolean; destination?: string }
        if (!status.configured) throw new Error('No live agent endpoint is configured. Set INJOFFICE_AGENT_PROPOSAL_URL on the local host, then reload this demo.')
        const context = await sessionInput.proposalContext!()
        if (cancelled) return
        setLiveContext(context)
        setLiveHost(status.destination || 'the configured agent endpoint')
        setLiveNotice('Review the shared context below before enabling the live request.')
        refreshTrace()
      } catch (reason) {
        if (!cancelled) {
          const message = reason instanceof Error ? reason.message : String(reason)
          setLiveNotice(message)
          if (proposalSource === 'mock') { setError(message); setState('error') }
        }
      }
    }
    void connect()
    return () => { cancelled = true; controller.abort() }
  }, [format, mode, proposalSource, sessionInput])

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
        if (proposalSource === 'mock') {
          if (!liveContext || !sessionInput.acceptProposal) throw new Error('The mock context is not ready. Reload the sample and try again.')
          const controller = new AbortController()
          proposalAbort.current = controller
          const { capabilities: sharedCapabilities, ...context } = liveContext
          const request = { request: prompt, context, capabilities: sharedCapabilities } as JsonObject
          const proposed = await requestMockAgentProposal(request, controller.signal)
          if (run !== generation.current) return
          setProposalExchange({ request, response: proposed })
          operations = sessionInput.acceptProposal(proposed)
        } else if (proposalSource === 'live') {
          if (!consent || !liveContext || !sessionInput.acceptProposal) throw new Error('Review the bounded context and explicitly consent before requesting a live proposal.')
          const controller = new AbortController()
          proposalAbort.current = controller
          const { capabilities: sharedCapabilities, ...context } = liveContext
          const response = await fetch('/api/agent/propose', {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, signal: controller.signal,
            body: JSON.stringify({ request: prompt, context, capabilities: sharedCapabilities, consent: true }),
          })
          if (!response.ok) throw new Error(`The proposal host could not complete the request (HTTP ${response.status}). No file change was committed.`)
          const proposed: unknown = await response.json()
          if (run !== generation.current) return
          operations = sessionInput.acceptProposal(proposed)
        } else operations = await sessionInput.propose(prompt)
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
  const boundary = mode === 'safe' && proposalSource === 'mock'
    ? `Simulated AI proposal · real ${fileType} preview, approval, file write, and verification · no language model or external service`
    : mode === 'safe' && proposalSource === 'live'
      ? `Real ${fileType} file · optional live proposal · only consented bounded context is shared · file write and verification stay local`
      : `Real ${fileType} file · local rule-based proposal through public tools · browser file write and exact-byte reopen · no model service or file upload`
  const requestHelp = format === 'xlsx'
    ? 'Try “Mark Mobile as On track” or “Mark Security as Ready”. Supports one workstream status edit at a time.'
    : format === 'pdf' ? 'Try “Rotate page 2 by 90 degrees” or “Rotate page 1 by 180 degrees”. Supports one page rotation at a time.'
    : 'Use Replace "exact current text" with "replacement text". Supports one uniquely identified editable text target at a time; copy the current text from the preview.'

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
          onChange={(id) => selectMode(id as AgentDemoMode)}
          options={[
            { id: 'safe', label: 'Supported change' },
            { id: 'refusal', label: 'Refusal proof' },
          ]}
        />
        <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={(!sessionInput && state !== 'error') || state === 'preparing' || state === 'committing' || safetyBusy || (sourceChanged && state === 'awaiting-approval') || (state === 'ready' && mode === 'safe' && proposalSource !== 'local' && (!liveContext || (proposalSource === 'live' && !consent)))} onClick={() => state === 'verified' || state === 'unverified' || state === 'error' ? reloadSample() : void prepare()}>{state === 'verified' || state === 'unverified' || state === 'error' ? 'Reload sample' : 'Run agent'}</DsButton>
        <span className="agent-demo__status" data-state={state} role="status" aria-live="polite">{status}</span>
      </div>

      <div className="agent-request">
        <label htmlFor="agent-prompt">Agent request</label>
        <textarea id="agent-prompt" data-agent-request maxLength={2000} readOnly={mode !== 'safe'} disabled={state === 'preparing' || state === 'committing' || state === 'verified' || state === 'unverified' || sourceChanged || safetyBusy} value={mode === 'safe' ? prompt : scenario.prompt} onChange={(event) => { reset(); setConsent(false); setPrompt(event.target.value) }} rows={2} />
        <small data-agent-boundary>{boundary}</small>
        {mode === 'safe' && <>
          <label className="agent-proposal-source">Proposal source <select data-agent-proposal-source value={proposalSource} disabled={state !== 'ready' || safetyBusy} onChange={(event) => { reset(); setProposalSource(event.target.value as 'mock' | 'local' | 'live') }}><option value="mock">Built-in mock agent (no LLM)</option><option value="local">Local rule-based proposer</option><option value="live">Live agent via configured host</option></select></label>
          <p className="agent-request-help">{requestHelp} The mock and local proposer are not language models.</p>
          {proposalSource === 'mock' && <div data-agent-mock className="agent-live-proposal">
            <p role="status">{liveNotice}</p>
            <p>The built-in mock endpoint returns a deterministic proposal from the inspected document. {import.meta.env.DEV ? 'It runs on this local demo server.' : 'On this static site, the endpoint response is simulated in your browser.'} No provider is contacted. Preview, approval, file editing, and verification are real.</p>
            {proposalExchange && <details data-agent-proposal-trace><summary>Mock proposal request and response</summary><pre className="ds-code" tabIndex={0}>{JSON.stringify(proposalExchange, null, 2)}</pre></details>}
          </div>}
          {proposalSource === 'live' && <div className="agent-live-proposal">
            <p role="status">{liveNotice}</p>
            {liveContext && <>
              <details data-agent-live-context><summary>Exact bounded context shared with {liveHost}</summary><pre className="ds-code" tabIndex={0}>{JSON.stringify(liveContext, null, 2)}</pre></details>
              <label className="ds-check"><input data-agent-live-consent type="checkbox" checked={consent} disabled={state !== 'ready'} onChange={(event) => setConsent(event.target.checked)} />Send this context and my request to {liveHost} for a proposal. This may incur charges from my configured provider. No Office file bytes or browser-stored API key are sent.</label>
            </>}
            <p>The endpoint can propose only the disclosed bounded edit. It cannot approve or commit. Use mock or local mode if you do not want to share document context.</p>
          </div>}
        </>}
      </div>

      <ol className="agent-tool-log" aria-label="Actual office.* tool calls">
        {AGENT_TOOL_METHODS.map((method) => (
          <li key={method} data-state={(sessionInput?.trace ? trace.some((entry) => entry.request.method === method && entry.response.ok) : toolLog.includes(method)) ? 'done' : 'waiting'}><code>{method}</code></li>
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
            <div><span>{verification?.ok ? 'Reopened output projection' : `Isolated ${preview ? 'preview' : 'source'}`}</span><h2 id="agent-artifact-title">{scenario.artifact.name}</h2></div>
            <dl className="ds-proof">
              <div><dt>Artifact</dt><dd>{scenario.artifact.artifactId}</dd></div>
              <div><dt>Revision</dt><dd>{receipt?.revision ?? inspection?.revision ?? scenario.artifact.revision}</dd></div>
            </dl>
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
            <header><div><span>Inspect</span><h2>Available operations</h2></div><b>{capabilities.length}</b></header>
            <div className="agent-capabilities">
              {capabilities.map((capability) => <span key={capability.operation} data-access={capability.destructive ? 'destructive' : capability.access}>{capability.operation}</span>)}
            </div>
          </section>

          <section className="ds-panel">
            <header><div><span>Plan</span><h2>{mode === 'safe' ? `Proposed ${fileType} change` : scenario.summary}</h2></div></header>
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
              <header><div><span>Validate</span><h2>Review before commit</h2></div></header>
              {state === 'verified' ? <DsChip tone="green">Applied</DsChip> : null}
              <label className="ds-check">
                <input type="checkbox" checked={approved} disabled={state !== 'awaiting-approval'} onChange={(event) => setApproved(event.target.checked)} />
                I reviewed this exact diff and approve one atomic commit.
              </label>
              <DsButton variant={state === 'verified' ? 'green' : 'filled'} className="workbench-button workbench-button--primary" disabled={!approved || state !== 'awaiting-approval' || safetyBusy} onClick={() => void commit()}>Commit approved change</DsButton>
              <p className="agent-approval-note">Approval is stored by the host for this exact change set. A tool argument saying “approved” cannot grant permission.</p>
            </section>
          )}

          <section className="agent-evidence ds-panel">
            <header><div><span>Commit</span><h2>{verification?.ok ? `${fileType} output verified` : 'Execution record'}</h2></div>{verification?.ok && <DsChip tone="green">Pass</DsChip>}</header>
            <dl className="ds-proof">
              <div><dt>Source revision</dt><dd><code>{inspection?.revision ?? '—'}</code></dd></div>
              <div><dt>Source fingerprint</dt><dd><code>{inspection?.fingerprint ?? '—'}</code></dd></div>
              <div><dt>Output revision</dt><dd><code>{receipt?.revision ?? '—'}</code></dd></div>
              <div><dt>Output fingerprint</dt><dd><code>{verification?.fingerprint ?? '—'}</code></dd></div>
            </dl>
            {verification && <ul>{verification.evidence.map((item) => <li key={item}>{item}</li>)}</ul>}
            <p>Verification above comes from the committed receipt, not a second <code>office.verify</code> call.</p>
            {state === 'verified' && verification?.ok && downloadURL && <a data-agent-download className="workbench-button ds-btn ds-btn--filled" href={downloadURL} download={`${scenario.artifact.name.replace(/\.[^.]+$/, '')}-approved.${format}`}>Download verified .{format}</a>}
          </section>
          {sessionInput?.stats && <section className="ds-panel agent-safety" aria-label="Safety scenarios">
            <header><h2>Try the safety boundaries</h2></header>
            <p>Source file writes: <span data-agent-native-writes>{nativeWrites}</span>. Isolated preview writes are not counted.</p>
            {sessionInput.simulateConcurrentEdit && <DsButton data-agent-concurrent-edit variant="outlined" disabled={state !== 'awaiting-approval' || sourceChanged || safetyBusy} onClick={() => void simulateConcurrentEdit()}>Simulate another editor changing the source</DsButton>}
            {sourceChanged && <p role="status">The source has changed. Try committing the reviewed plan: its stale revision must be rejected.</p>}
            {sessionInput.setVerificationFailure && <><label className="ds-check"><input data-agent-fail-verification type="checkbox" checked={failVerification} disabled={state !== 'awaiting-approval' || safetyBusy} onChange={(event) => setFailVerification(event.target.checked)} />Inject a verification-read failure after the write</label>
            <p>This explicitly simulates a readback failure; it does not corrupt the sample file.</p></>}
            <DsButton data-agent-retry variant="outlined" disabled={state !== 'verified' || safetyBusy} onClick={() => void retryCommit()}>Retry the same commit</DsButton>
            {retryResult && <p role="status">{retryResult}</p>}
          </section>}
          {error && <p className="tool-error" role="alert">{error}</p>}
        </aside>
      </div>
      <details data-agent-trace className="agent-dispatch-trace">
        <summary>Actual tool requests and results ({trace.length})</summary>
        <p>Requests pass through <code>createAgentToolDispatcher</code> and the public {fileType} adapter. The final commit result contains post-write verification. Read and preview results are bounded document projections.</p>
        {trace.map((entry) => <details key={entry.request.requestId}>
          <summary>{entry.request.method} · {entry.response.ok ? 'Completed' : 'Refused or failed'}</summary>
          <pre className="ds-code" tabIndex={0}>{JSON.stringify(entry, null, 2)}</pre>
        </details>)}
      </details>
    </section>
    </div>
  )
}
