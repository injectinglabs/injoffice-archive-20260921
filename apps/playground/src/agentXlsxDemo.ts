import { createXlsxAgentAdapter, type XlsxAgentArtifact } from '@injoffice/agent-office/xlsx'
import {
  AGENT_TOOLS_PROTOCOL, AGENT_TOOLS_PROTOCOL_VERSION, createAgentSession, createAgentToolDispatcher,
  type AgentArtifactView, type AgentBoundedResult, type AgentCapability, type AgentChangeSetEnvelope,
  type AgentCommitResult, type AgentToolCall, type AgentToolCallResult, type AgentToolMethod,
  type AgentValidationReport, type JsonObject,
} from '@injoffice/agent-tools'
import { adaptWorkbookMutationBatchV1 } from '@injoffice/xlsx-wasm'
import type { WorkbookMutationBatchV1 } from '@injoffice/sheets/browser'
import type { AgentDemoSession } from './agentDemoRuntime'
import type { AgentDemoArtifact, AgentDemoMode, AgentDemoOperation, AgentDemoScenario } from './agentDemoScenario'
import { buildCellMutation, displayCellValue, editableTargets, type NativeWorkbook } from './nativeRoundTrip'
import { createBrowserXlsxRoundTripRuntime, type XlsxRoundTripRuntime } from './xlsxRoundTripRuntime'

export const AGENT_XLSX_NAME = 'launch-readiness-plan.xlsx'
const MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const BOUNDS = { maxItems: 100, maxBytes: 64_000 }
export type NativeAgentTrace = { request: AgentToolCall; response: AgentToolCallResult }

function projection(workbook: NativeWorkbook, operation?: AgentDemoOperation): Record<string, unknown> {
  const sheet = workbook.sheets.find((item) => item.id === operation?.input.sheetId) ?? workbook.sheets[0]
  const cell = (row: number, column: number) => displayCellValue(sheet.cells.find((item) => item.row === row && item.column === column))
  const target = operation?.input.cell as { row: number; column: number } | undefined
  const columns = Math.min(12, Math.max(5, ...sheet.cells.map((item) => item.column)) + 1)
  const rows = Math.min(100, Math.max(7, ...sheet.cells.map((item) => item.row)))
  return { sheet: { id: sheet.id, name: sheet.name }, headers: Array.from({ length: columns }, (_, column) => cell(0, column)),
    rows: Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) => cell(row + 1, column))),
    ...(target ? { changedCell: { row: target.row - 1, column: target.column } } : {}) }
}

/** Local host for the shipped XLSX adapter and JSON dispatcher. No model or remote service is called. */
export async function createNativeAgentSessionInput(
  mode: AgentDemoMode,
  runtime: XlsxRoundTripRuntime = createBrowserXlsxRoundTripRuntime(),
  fetcher: typeof fetch = globalThis.fetch.bind(globalThis),
) {
  let disposed = false
  let verifiedBytes: Uint8Array | null = null
  let nativeWrites = 0
  let failVerification = false
  let postCommitReads: number | null = null
  const approvals = new Map<string, string>()
  const traces: NativeAgentTrace[] = []
  const plans = new Map<string, AgentChangeSetEnvelope>()
  const previews = new Map<string, NativeWorkbook>()
  const proposalRevisions = new Map<string, string>()
  const proposalKey = (operations: readonly AgentDemoOperation[]) => JSON.stringify(operations.map((item) => [item.op, item.input.sheetId, item.input.cell, item.input.value]))
  const assertOpen = () => { if (disposed) throw new Error('This workbook session is closed. Reload the sample to continue.') }
  const dispose = () => { if (!disposed) { disposed = true; runtime.terminate(); verifiedBytes = null; approvals.clear(); previews.clear() } }
  try {
    const response = await fetcher(`${import.meta.env.BASE_URL}native-fixture/${AGENT_XLSX_NAME}`)
    if (!response.ok) throw new Error(`The sample workbook could not be loaded (HTTP ${response.status}).`)
    let bytes = new Uint8Array(await response.arrayBuffer())
    let workbook = (await runtime.extract(bytes.slice())).workbook
    let selectedOperation: AgentDemoOperation | undefined
    const artifactView = (value = workbook): AgentDemoArtifact => ({ artifactId: 'launch-readiness-agent-sample', format: 'xlsx',
      name: AGENT_XLSX_NAME, revision: value.source.package_sha256, content: projection(value, selectedOperation) })
    // This tab owns persistence. Atomic writes and the idempotency ledger are scoped to this session, not durable across reloads.
    let mutationQueue = Promise.resolve()
    const serialize = <T>(action: () => Promise<T>): Promise<T> => {
      const result = mutationQueue.then(action)
      mutationQueue = result.then(() => undefined, () => undefined)
      return result
    }
    const committedBatches = new Map<string, { signature: string; workbook: NativeWorkbook }>()
    const applyCopy = async (batch: WorkbookMutationBatchV1) => {
      assertOpen()
      const source = workbook
      if (batch.expected_revision !== source.source.package_sha256) throw new Error('The source revision changed before the native write.')
      const transaction = adaptWorkbookMutationBatchV1(source, batch)
      const applied = await runtime.apply({ original: bytes.slice(), workbook: structuredClone(source), transaction, sourceName: AGENT_XLSX_NAME })
      const replacement = applied.bytes.slice()
      const reopened = (await runtime.extract(replacement.slice())).workbook
      assertOpen()
      return { bytes: replacement, workbook: reopened, sourceRevision: source.source.package_sha256 }
    }
    const host: XlsxAgentArtifact = {
      kind: 'injoffice.xlsx', artifactId: 'launch-readiness-agent-sample',
      async snapshot() {
        assertOpen()
        // Core first reopens the replacement to check its identity, then the public adapter independently reads it for verification.
        // The labelled fault affects only that latter proof read, never the successful write or replacement identity guard.
        if (postCommitReads !== null) {
          const stage = postCommitReads++
          if (stage === 1) {
            postCommitReads = null
            if (failVerification) throw new Error('Demo fault: post-commit verification read unavailable. The write already succeeded.')
          }
        }
        const reopened = (await runtime.extract(bytes.slice())).workbook
        assertOpen()
        return { workbook: reopened }
      },
      async preview(batch) {
        const candidate = await applyCopy(batch)
        previews.set(candidate.workbook.source.package_sha256, candidate.workbook)
        return { workbook: candidate.workbook }
      },
      apply(batch) {
        return serialize(async () => {
          assertOpen()
          const signature = JSON.stringify(batch)
          const retained = committedBatches.get(batch.batch_id)
          if (retained) {
            if (retained.signature !== signature) throw new Error('A different native batch already used this idempotency key.')
            return { snapshot: { workbook: structuredClone(retained.workbook) }, deduplicated: true }
          }
          const candidate = await applyCopy(batch)
          if (workbook.source.package_sha256 !== candidate.sourceRevision) throw new Error('The source revision changed during the native write.')
          bytes = candidate.bytes; workbook = candidate.workbook; nativeWrites += 1; verifiedBytes = null
          postCommitReads = 0
          committedBatches.set(batch.batch_id, { signature, workbook: structuredClone(workbook) })
          return { snapshot: { workbook: structuredClone(workbook) }, receipt: { nativeWrites }, deduplicated: false }
        })
      },
    }
    const core = await createAgentSession({ artifact: host, adapter: createXlsxAgentAdapter(),
      actor: { id: 'playground-local-host', kind: 'agent' },
      confirmDestructive: ({ changeSet, identity }) => !!changeSet && approvals.get(changeSet.changeSetId) === JSON.stringify(changeSet)
        && changeSet.baseRevision === identity.revision && changeSet.baseFingerprint === identity.fingerprint,
    })
    const dispatcher = createAgentToolDispatcher(core)
    let requestNumber = 0
    const call = async <T>(method: AgentToolMethod, params: JsonObject = {}): Promise<T> => {
      assertOpen()
      const request: AgentToolCall = { protocol: AGENT_TOOLS_PROTOCOL, protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
        requestId: `demo-tool-${++requestNumber}`, method, params }
      const result = await dispatcher.dispatch(request)
      assertOpen()
      traces.push(structuredClone({ request, response: result }))
      if (!result.ok) throw new Error(`${result.error.code}: ${result.error.message}${result.error.issues?.length ? ` ${result.error.issues.map((item) => `${item.code}: ${item.message}`).join('; ')}` : ''}`)
      return result.result as T
    }
    const readSheet = async (sheetId: string, revision?: string) => {
      const cells: Array<NativeWorkbook['sheets'][number]['cells'][number]> = []
      let cursor: string | undefined
      do {
        const result = await call<AgentBoundedResult>('office.read', { query: { sheetId }, ...BOUNDS, ...(cursor ? { cursor } : {}) })
        if (revision && result.identity.revision !== revision) throw new Error('The document changed during discovery. Read it again before proposing.')
        cells.push(...(result.data as unknown as { cells: NativeWorkbook['sheets'][number]['cells'] }).cells)
        if (result.truncated && !result.nextCursor) throw new Error('The workbook content exceeds the bounded discovery budget.')
        cursor = result.nextCursor
        if (cells.length > 2_000) throw new Error('The local proposal supports at most 2,000 inspected cells per sheet.')
      } while (cursor)
      return cells
    }
    const statusValues = ['Ready', 'On track', 'At risk', 'Review', 'Blocked']
    type DisclosedTarget = { sheetId: string; sheetName: string; row: number; column: number; ref: string; workstream: string; currentValue: string }
    let disclosed: { revision: string; targets: DisclosedTarget[] } | undefined
    const proposalContext = async (): Promise<JsonObject> => {
      const capabilities = await call<{ capabilities: AgentCapability[] }>('office.capabilities')
      const inspection = await call<AgentBoundedResult>('office.inspect', BOUNDS)
      if (inspection.truncated) throw new Error('The workbook structure exceeds the bounded proposal context.')
      const sheets = (inspection.data as unknown as { sheets: Array<{ id: string; name: string; editable: boolean }> }).sheets
      if (sheets.length > 8) throw new Error('The local proposal bridge supports at most eight sample sheets.')
      const targets: DisclosedTarget[] = []
      const contextSheets: JsonObject[] = []
      for (const sheet of sheets) {
        if (!sheet.editable) continue
        const cells = await readSheet(sheet.id, inspection.identity.revision)
        contextSheets.push({ id: sheet.id, name: sheet.name, cells: cells.map((cell) => ({ ref: cell.ref, row: cell.row, column: cell.column, value: displayCellValue(cell), editable: cell.editable })) })
        const header = cells.find((cell) => displayCellValue(cell).trim().toLowerCase() === 'workstream')
        const status = header && cells.find((cell) => cell.row === header.row && displayCellValue(cell).trim().toLowerCase() === 'status')
        if (!header || !status) continue
        for (const name of cells.filter((cell) => cell.column === header.column && cell.row > header.row)) {
          const target = cells.find((cell) => cell.row === name.row && cell.column === status.column)
          if (target?.editable && !target.formula) targets.push({ sheetId: sheet.id, sheetName: sheet.name, row: target.row, column: target.column,
            ref: target.ref, workstream: displayCellValue(name), currentValue: displayCellValue(target) })
        }
      }
      disclosed = { revision: inspection.identity.revision, targets }
      return { identity: { artifactId: inspection.identity.artifactId, revision: inspection.identity.revision, fingerprint: inspection.identity.fingerprint },
        capabilities: capabilities.capabilities.filter((item) => item.name === 'xlsx.cell.set_value') as unknown as JsonObject[],
        workbook: { sheets: contextSheets }, constraints: { maxOperations: 1, allowedValues: statusValues, allowedTargets: targets.map((target) => ({ ...target })) } }
    }
    const acceptProposal = (value: unknown): AgentDemoOperation[] => {
      assertOpen()
      const object = (input: unknown): input is Record<string, unknown> => typeof input === 'object' && input !== null && !Array.isArray(input)
      if (!disclosed) throw new Error('Read and explicitly disclose the proposal context before accepting a remote proposal.')
      if (!object(value) || Object.keys(value).some((key) => key !== 'operations') || !Array.isArray(value.operations) || value.operations.length !== 1) throw new Error('A proposal must contain exactly one operations entry and no other top-level fields.')
      const operation = value.operations[0]
      if (!object(operation) || Object.keys(operation).some((key) => !['name', 'input', 'operationId'].includes(key)) || operation.name !== 'xlsx.cell.set_value' || !object(operation.input)) throw new Error('Only a bounded xlsx.cell.set_value proposal is accepted.')
      const input = operation.input
      if (Object.keys(input).some((key) => !['sheetId', 'cell', 'value'].includes(key)) || !object(input.cell) || Object.keys(input.cell).some((key) => !['row', 'column'].includes(key)) || typeof input.value !== 'string' || !statusValues.includes(input.value)) throw new Error('The proposal must use an allowed status value and exact cell coordinates.')
      const cell = input.cell
      const target = disclosed.targets.find((item) => item.sheetId === input.sheetId && item.row === cell.row && item.column === cell.column)
      if (!target) throw new Error('The proposed cell is outside the disclosed editable status targets.')
      if (operation.operationId !== undefined && (typeof operation.operationId !== 'string' || !/^[A-Za-z0-9._:-]{1,128}$/.test(operation.operationId))) throw new Error('The operation id must be a short identifier.')
      const accepted: AgentDemoOperation[] = [{ operationId: typeof operation.operationId === 'string' ? operation.operationId : `agent-status-${target.row}-${target.column}`,
        op: 'xlsx.cell.set_value', target: `${target.sheetName}!${target.ref}`, value: input.value,
        input: { sheetId: target.sheetId, cell: { row: target.row, column: target.column }, value: input.value } }]
      selectedOperation = accepted[0]
      proposalRevisions.set(proposalKey(accepted), disclosed.revision)
      return accepted
    }
    const propose = async (request: string): Promise<AgentDemoOperation[]> => {
      const capabilities = await call<{ capabilities: AgentCapability[] }>('office.capabilities')
      if (!capabilities.capabilities.some((item) => item.name === 'xlsx.cell.set_value')) throw new Error('This workbook does not support literal cell writes.')
      const inspection = await call<AgentBoundedResult>('office.inspect', BOUNDS)
      if (/\bmacro\b/i.test(request)) return [{ operationId: 'agent-macro-refusal', op: 'macro.execute', target: 'workbook', input: {} }]
      const match = /^\s*mark\s+(.+?)\s+as\s+(ready|on track|at risk|review|blocked)(?:[.!]|\s+after\b.*)?\s*$/i.exec(request)
      if (!match) throw new Error('Try “Mark Security as Ready” or “Mark Mobile as On track”. This local planner supports one workstream status change, not arbitrary language.')
      const value = statusValues.find((item) => item.toLowerCase() === match[2].toLowerCase())!
      const matches: AgentDemoOperation[] = []
      const sheets = (inspection.data as unknown as { sheets: Array<{ id: string; name: string; editable: boolean }> }).sheets
      for (const sheet of sheets) {
        if (!sheet.editable) continue
        const cells = await readSheet(sheet.id, inspection.identity.revision)
        const header = cells.find((item) => displayCellValue(item).trim().toLowerCase() === 'workstream')
        const status = header && cells.find((item) => item.row === header.row && displayCellValue(item).trim().toLowerCase() === 'status')
        if (!header || !status) continue
        for (const name of cells.filter((item) => item.column === header.column && item.row > header.row && displayCellValue(item).trim().toLowerCase() === match[1].trim().toLowerCase())) {
          const target = cells.find((item) => item.row === name.row && item.column === status.column)
          if (!target?.editable || target.formula) throw new Error('The discovered status cell is not an editable literal.')
          matches.push({ operationId: `agent-status-${name.row}-${status.column}`, op: 'xlsx.cell.set_value', target: `${sheet.name}!${target.ref}`, value,
            input: { sheetId: sheet.id, cell: { row: name.row, column: status.column }, value } })
        }
      }
      if (matches.length !== 1) throw new Error(matches.length ? 'That workstream is ambiguous across the workbook. No change was proposed.' : `No unique editable status was found for “${match[1].trim()}”.`)
      selectedOperation = matches[0]
      proposalRevisions.set(proposalKey(matches), inspection.identity.revision)
      return matches
    }
    const session: AgentDemoSession = {
      async capabilities() {
        const result = await call<{ capabilities: AgentCapability[] }>('office.capabilities')
        return { format: 'xlsx', operations: result.capabilities.map((item) => ({ operation: item.name, access: 'write', destructive: item.destructive })) }
      },
      async inspect({ selection, maxItems, maxBytes }) {
        const result = await call<AgentBoundedResult>('office.inspect', { maxItems, maxBytes })
        return { artifactId: result.identity.artifactId, revision: result.identity.revision, fingerprint: result.identity.fingerprint,
          selection, content: projection(workbook, selectedOperation), truncated: result.truncated }
      },
      async plan(operations, { expectedRevision }) {
        const proposalRevision = proposalRevisions.get(proposalKey(operations))
        if (proposalRevision && proposalRevision !== expectedRevision) throw new Error('The document changed after proposal discovery. Prepare a new proposal.')
        const envelope = await call<AgentChangeSetEnvelope>('office.plan', { expectedRevision,
          operations: operations.map((item) => ({ operationId: item.operationId, name: item.op, input: item.input as JsonObject })) })
        plans.set(envelope.changeSetId, envelope)
        const params = { changeSetId: envelope.changeSetId }
        const before = structuredClone(workbook)
        const selected = operations[0]
        return { id: envelope.changeSetId, expectedRevision: envelope.baseRevision, operations: structuredClone(operations),
          async validate() {
            const result = await call<AgentValidationReport>('office.validate', params)
            return { ok: result.valid, issues: result.issues.map((item) => ({ code: item.code.toLowerCase(), path: item.path ?? '', message: item.message })) }
          },
          async preview() {
            const result = await call<AgentArtifactView>('office.preview', params)
            const failedProof = result.issues.filter((item) => item.severity === 'error')
            if (failedProof.length) throw new Error(`Native preview failed: ${failedProof.map((item) => `${item.code}: ${item.message}`).join('; ')}. No source write was committed.`)
            const candidate = previews.get(String((result.data as JsonObject).resultingRevision)) ?? before
            return { artifact: { ...artifactView(candidate), content: projection(candidate, selected) },
              summary: result.issues.length ? result.issues.map((item) => item.message).join('; ') : 'Native preview applied to a copy and reopened; source unchanged.',
              evidence: result.evidence.map((item) => item.description ?? item.kind) }
          },
          async diff() {
            const result = await call<AgentArtifactView>('office.diff', params)
            return { changes: result.issues.length ? [] : operations.map((operation) => {
              const cell = operation.input.cell as { row?: number; column?: number } | undefined
              const target = before.sheets.find((sheet) => sheet.id === operation.input.sheetId)?.cells.find((item) => item.row === cell?.row && item.column === cell?.column)
              return { target: operation.target, before: displayCellValue(target), after: String(operation.value ?? '') }
            }) }
          },
          async commit(input) {
            if (input.expectedRevision !== envelope.baseRevision) throw new Error('Commit revision does not match the reviewed change set.')
            const result = await call<AgentCommitResult>('office.commit', { ...params, idempotencyKey: input.idempotencyKey, confirmation: input.confirmation })
            // Release only bytes matching this successful receipt and the public adapter’s fresh native readback.
            if (result.verification.verified && result.identity.fingerprint === workbook.source.package_sha256 && result.identity.revision === workbook.source.package_sha256) verifiedBytes = bytes.slice()
            else verifiedBytes = null
            return { artifactId: result.identity.artifactId, previousRevision: envelope.baseRevision, revision: result.identity.revision,
              fingerprint: result.identity.fingerprint, operationIds: envelope.operations.map((item) => item.operationId),
              evidence: result.evidence?.map((item) => item.description ?? item.kind) ?? [], verification: result.verification, content: projection(workbook, selected) }
          },
          async verify(receipt) {
            // This is receipt presentation, not office.verify (which checks a PLANNED change set).
            return { ok: receipt.verification.verified, revision: receipt.revision, fingerprint: receipt.fingerprint,
              evidence: receipt.verification.checks.map((item) => item.message ?? item.name) }
          },
        }
      },
    }
    const prompt = mode === 'safe' ? 'Mark Security as Ready' : 'Execute an unsupported workbook macro.'
    const operations = await propose(prompt)
    const scenario: AgentDemoScenario = { artifact: artifactView(), operations, prompt,
      summary: mode === 'safe' ? 'Discover a status cell, review a native preview, approve its exact change set, and verify the saved XLSX.' : 'Refuse macro execution before any file write.' }
    return { scenario, operations, session, dispose, propose, proposalContext, acceptProposal,
      async approve(changeSetId: string) {
        assertOpen()
        const envelope = plans.get(changeSetId)
        if (!envelope) throw new Error('Only an exact change set retained by this host session can be approved.')
        const current = await call<AgentBoundedResult>('office.inspect', BOUNDS)
        if (current.identity.revision !== envelope.baseRevision || current.identity.fingerprint !== envelope.baseFingerprint) throw new Error('The document changed. Prepare and review a new proposal before approval.')
        approvals.set(changeSetId, JSON.stringify(envelope))
      },
      async simulateConcurrentEdit() {
        await serialize(async () => {
          assertOpen()
          const target = editableTargets(workbook).find((item) => item.row > 0 && item.column === 1 && !item.formula)
          if (!target) throw new Error('The sample has no editable owner cell for the concurrent-edit scenario.')
          const batch = buildCellMutation(workbook, target, `${displayCellValue(target)} (reviewed)`, 'host-concurrent-edit')
          const candidate = await applyCopy(batch)
          bytes = candidate.bytes; workbook = candidate.workbook; nativeWrites += 1; verifiedBytes = null; approvals.clear(); postCommitReads = null
        })
      },
      setVerificationFailure(enabled: boolean) { assertOpen(); failVerification = enabled },
      trace: () => structuredClone(traces), stats: () => ({ nativeWrites }),
      download: () => !disposed && verifiedBytes ? new Blob([new Uint8Array(verifiedBytes).buffer], { type: MEDIA_TYPE }) : null,
    }
  } catch (error) { dispose(); throw error }
}
