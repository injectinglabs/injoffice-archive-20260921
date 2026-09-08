import type { AgentArtifactAdapter, AgentArtifactIdentity, AgentCapability, AgentIssue, AgentOperation, JsonValue } from '@injoffice/agent-tools'
import { adaptWorkbookMutationBatchV1 } from '@injoffice/xlsx-wasm'
import { createAgentDemoSession } from './agentDemoRuntime'
import type { AgentDemoArtifact, AgentDemoMode, AgentDemoScenario } from './agentDemoScenario'
import { buildCellMutation, displayCellValue, editableTargets, type NativeWorkbook } from './nativeRoundTrip'
import { createBrowserXlsxRoundTripRuntime, type XlsxRoundTripRuntime } from './xlsxRoundTripRuntime'

export const AGENT_XLSX_NAME = 'launch-readiness-plan.xlsx'
const MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const TARGET_REF = 'C5'
const VALUE = 'Ready'
const proof = (kind: string, description: string) => ({ kind, description })

function projection(workbook: NativeWorkbook): Record<string, unknown> {
  const sheet = workbook.sheets[0]
  const cell = (row: number, column: number) => displayCellValue(sheet.cells.find((item) => item.row === row && item.column === column))
  return {
    sheet: { id: sheet.id, name: sheet.name },
    headers: Array.from({ length: 6 }, (_, column) => cell(0, column)),
    rows: Array.from({ length: 7 }, (_, row) => Array.from({ length: 6 }, (_, column) => cell(row + 1, column))),
    changedCell: { row: 3, column: 2 },
  }
}

/** A deliberately bounded real-file adapter: one reviewed literal in the bundled workbook. */
export async function createNativeAgentSessionInput(
  mode: AgentDemoMode,
  runtime: XlsxRoundTripRuntime = createBrowserXlsxRoundTripRuntime(),
  fetcher: typeof fetch = globalThis.fetch.bind(globalThis),
) {
  const records = new WeakMap<AgentDemoArtifact, { bytes: Uint8Array; workbook: NativeWorkbook }>()
  let verifiedBytes: Uint8Array | null = null
  let disposed = false
  const dispose = () => { if (!disposed) { disposed = true; runtime.terminate(); verifiedBytes = null } }
  try {
    const response = await fetcher(`${import.meta.env.BASE_URL}native-fixture/${AGENT_XLSX_NAME}`)
    if (!response.ok) throw new Error(`The sample workbook could not be loaded (HTTP ${response.status}).`)
    const sourceBytes = new Uint8Array(await response.arrayBuffer())
    const { workbook } = await runtime.extract(sourceBytes.slice())
    const target = editableTargets(workbook).find((cell) => cell.sheetId === workbook.sheets[0].id && cell.ref === TARGET_REF)
    const workstream = workbook.sheets[0]?.cells.find((cell) => cell.ref === 'A5')
    if (!target || displayCellValue(target) !== 'Review' || displayCellValue(workstream) !== 'Security') throw new Error('The sample no longer contains the expected editable Security status at C5.')
    const artifact: AgentDemoArtifact = {
      artifactId: 'launch-readiness-agent-sample', format: 'xlsx', name: AGENT_XLSX_NAME,
      revision: workbook.revision, content: projection(workbook),
    }
    records.set(artifact, { bytes: sourceBytes, workbook })
    const record = (item: AgentDemoArtifact) => {
      if (disposed) throw new Error('This workbook session is closed. Reload the sample to continue.')
      const data = records.get(item)
      if (!data) throw new Error('Unknown workbook artifact.')
      return data
    }
    const identity = (item: AgentDemoArtifact): AgentArtifactIdentity => ({
      artifactId: item.artifactId, format: 'xlsx', mediaType: MEDIA_TYPE,
      revision: record(item).workbook.revision, fingerprint: record(item).workbook.source.package_sha256,
    })
    const issues = (operations: readonly AgentOperation[]): AgentIssue[] => operations.flatMap((operation) => {
      const cell = operation.input.cell as { row?: unknown; column?: unknown } | undefined
      const valid = operation.name === 'xlsx.cell.set_value' && operation.input.sheetId === target.sheetId &&
        cell?.row === target.row && cell?.column === target.column && operation.input.value === VALUE
      return valid ? [] : [{ severity: 'refusal' as const, code: 'UNSUPPORTED_OPERATION', operationId: operation.operationId,
        path: `/operations/${operation.operationId}`, message: 'This real-file example only supports the reviewed Security status change at C5. No output was written.' }]
    })
    const adapter: AgentArtifactAdapter<AgentDemoArtifact> = {
      id: 'playground-native-xlsx', format: 'xlsx',
      supports: (item): item is AgentDemoArtifact => typeof item === 'object' && item !== null && records.has(item as AgentDemoArtifact),
      async identity({ artifact: item }) { return identity(item) },
      async capabilities(): Promise<AgentCapability[]> { return [
        { name: 'xlsx.inspect', description: 'Read the bundled workbook through the native XLSX engine.', inputSchema: { type: 'object', additionalProperties: true } },
        { name: 'xlsx.cell.set_value', description: 'Set the Security status at C5 to Ready in real XLSX bytes.', requiresConfirmation: true,
          inputSchema: { type: 'object', properties: { sheetId: { type: 'string' }, cell: { type: 'object' }, value: { type: 'string' } }, required: ['sheetId', 'cell', 'value'], additionalProperties: false } },
      ] },
      async inspect({ artifact: item }) { return { data: projection(record(item).workbook) as JsonValue, itemCount: 8, truncated: false, evidence: [proof('native-extraction', 'Read eight rows from the bundled XLSX bytes in a browser Worker.')] } },
      async read({ artifact: item }) { return { data: projection(record(item).workbook) as JsonValue, itemCount: 8, truncated: false } },
      async validate({ changeSet }) {
        if (changeSet.operations.length !== 1) return [{ severity: 'refusal', code: 'ONE_OPERATION_ONLY', message: 'Review exactly one status change per sample.' }]
        return issues(changeSet.operations)
      },
      async preview({ artifact: item, changeSet }) {
        const invalid = issues(changeSet.operations)
        const preview = structuredClone(item)
        if (!invalid.length) (preview.content.rows as string[][])[3][2] = VALUE
        return { data: preview as unknown as JsonValue, issues: invalid, evidence: [proof('projected-preview', 'Projected cell preview only; the original XLSX bytes have not been changed.')] }
      },
      async diff({ artifact: item, changeSet }) {
        return { data: { changes: issues(changeSet.operations).length ? [] : [{ target: `${target.sheetName}!${TARGET_REF}`, before: displayCellValue(editableTargets(record(item).workbook).find((cell) => cell.ref === TARGET_REF && cell.sheetName === target.sheetName)), after: VALUE }] }, issues: issues(changeSet.operations), evidence: [] }
      },
      async commit({ artifact: item, changeSet, expectedRevision, expectedFingerprint }) {
        const current = identity(item)
        if (current.revision !== expectedRevision || current.fingerprint !== expectedFingerprint) throw new Error('The source revision changed before commit.')
        if (changeSet.operations.length !== 1 || issues(changeSet.operations).length) throw new Error('The proposed operation was refused before write.')
        const source = record(item)
        const batch = buildCellMutation(source.workbook, target, VALUE, changeSet.operations[0].operationId)
        const transaction = adaptWorkbookMutationBatchV1(source.workbook, batch)
        const applied = await runtime.apply({ original: source.bytes.slice(), workbook: source.workbook, transaction, sourceName: AGENT_XLSX_NAME })
        const bytes = applied.bytes.slice()
        const reopened = await runtime.extract(bytes.slice())
        const next: AgentDemoArtifact = { ...item, revision: reopened.workbook.revision, content: projection(reopened.workbook) }
        records.set(next, { bytes, workbook: reopened.workbook })
        return { artifact: next, identity: identity(next), data: next.content as JsonValue, evidence: [proof('native-write', 'Applied the approved mutation to real XLSX bytes and reopened the exact replacement package.')] }
      },
      async verify({ artifact: item, stage, commit }) {
        if (stage === 'planned') return { verified: true, issues: [], checks: [{ name: 'source-isolated', passed: true, message: 'The original XLSX bytes are unchanged; approval is required before write.' }] }
        const output = record(item)
        // Independently reopen the exact downloadable bytes, not the preview object.
        const reopened = (await runtime.extract(output.bytes.slice())).workbook
        const result = editableTargets(reopened).find((cell) => cell.sheetName === target.sheetName && cell.ref === TARGET_REF)
        const checks = [
          { name: 'native-cell-readback', passed: displayCellValue(result) === VALUE, message: `Native readback of ${target.sheetName}!${TARGET_REF}: ${displayCellValue(result)}.` },
          { name: 'exact-output-identity', passed: reopened.source.package_sha256 === commit?.identity.fingerprint && reopened.revision === commit?.identity.revision, message: 'Reopened XLSX revision and package SHA-256 match the commit receipt.' },
          { name: 'revision-advanced', passed: reopened.source.package_sha256 !== workbook.source.package_sha256, message: 'The replacement differs from the source package.' },
        ]
        const verified = checks.every((check) => check.passed)
        verifiedBytes = verified ? output.bytes.slice() : null
        return { verified, checks, issues: [] }
      },
    }
    const safeOperation = { operationId: 'agent-security-status', op: 'xlsx.cell.set_value', target: `${target.sheetName}!${TARGET_REF}`, value: VALUE,
      input: { sheetId: target.sheetId, cell: { row: target.row, column: target.column }, value: VALUE } }
    const operations = mode === 'safe' ? [safeOperation] : [{ operationId: 'agent-macro-refusal', op: 'macro.execute', target: 'workbook', input: {} }]
    const scenario: AgentDemoScenario = { artifact, operations,
      prompt: mode === 'safe' ? 'Mark Security as Ready after the approved pen-test review.' : 'Execute an unsupported workbook macro.',
      summary: mode === 'safe' ? `Change ${target.sheetName}!C5 from Review to Ready, then reopen the saved XLSX.` : 'Refuse macro execution before any file write.',
    }
    return { scenario, operations, session: createAgentDemoSession(artifact, adapter), dispose,
      download: () => verifiedBytes ? new Blob([new Uint8Array(verifiedBytes).buffer], { type: MEDIA_TYPE }) : null }
  } catch (error) { dispose(); throw error }
}
