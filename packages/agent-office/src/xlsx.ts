import {
  WORKBOOK_MUTATION_PROTOCOL, WORKBOOK_MUTATION_VERSION, decodeWorkbookMutationBatch, validateNativeWorkbookV1, validateNativeWorkbookV2,
  type NativeWorkbookV1, type NativeWorkbookV2, type StyleDelta, type SupportedWorkbookMutation, type WorkbookMutationBatchV1,
} from '@injoffice/sheets/browser'
import type { AgentArtifactAdapter, AgentArtifactIdentity, AgentCapability, AgentIssue, AgentOperation, JsonObject, JsonValue } from '@injoffice/agent-tools'
import { assertFresh, boundedObject, cursorOffset, errorIssue, toJson } from './hash.js'

export type XlsxNativeWorkbook = NativeWorkbookV1 | NativeWorkbookV2
export interface XlsxAgentSnapshot { workbook: XlsxNativeWorkbook }
export interface XlsxAgentArtifact {
  readonly kind: 'injoffice.xlsx'; readonly artifactId: string
  snapshot(options?: { signal?: AbortSignal }): Promise<XlsxAgentSnapshot>
  /** Must atomically enforce expected_revision and durably deduplicate batch_id. */
  apply(batch: WorkbookMutationBatchV1, options?: { signal?: AbortSignal }): Promise<{ snapshot: XlsxAgentSnapshot; receipt?: JsonValue; deduplicated?: boolean }>
  preview?(batch: WorkbookMutationBatchV1, options?: { signal?: AbortSignal }): Promise<XlsxAgentSnapshot>
}

const object = (properties: JsonObject, required: string[]): JsonObject => ({ type: 'object', properties, required, additionalProperties: false })
const cellProperties: JsonObject = { row: { type: 'integer', minimum: 0, maximum: 1_048_575 }, column: { type: 'integer', minimum: 0, maximum: 16_383 } }
const cell = object(cellProperties, ['row', 'column'])
const range = object({ ...cellProperties, endRow: { type: 'integer', minimum: 0, maximum: 1_048_575 }, endColumn: { type: 'integer', minimum: 0, maximum: 16_383 } }, ['row', 'column', 'endRow', 'endColumn'])
const style = object({
  number_format: { type: 'string', maxLength: 255 }, font_name: { type: 'string', maxLength: 255 },
  font_size_points: { type: 'number', minimum: 1, maximum: 409 }, bold: { type: 'boolean' }, italic: { type: 'boolean' },
  font_color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' }, fill_color: { type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' },
  horizontal_alignment: { type: 'string', enum: ['general', 'left', 'center', 'right'] }, vertical_alignment: { type: 'string', enum: ['top', 'middle', 'bottom'] },
  wrap_text: { type: 'boolean' },
}, [])
const capability = (name: string, description: string, properties: JsonObject, required: string[]): AgentCapability => ({
  name, description, destructive: true, requiresConfirmation: true,
  inputSchema: { type: 'object', properties: { sheetId: { type: 'string', minLength: 1, maxLength: 256 }, ...properties }, required: ['sheetId', ...required], additionalProperties: false },
  metadata: { execution: 'host', wholeFileWrite: true },
})
const CAPABILITIES: readonly AgentCapability[] = [
  capability('xlsx.cell.set_value', 'Set one literal cell value.', { cell, value: { type: ['string', 'number', 'boolean'] } }, ['cell', 'value']),
  capability('xlsx.cell.clear_value', 'Clear one literal cell value.', { cell }, ['cell']),
  capability('xlsx.cell.set_formula', 'Set one A1 formula including the leading equals sign.', { cell, formula: { type: 'string', minLength: 2, maxLength: 8192 } }, ['cell', 'formula']),
  capability('xlsx.cell.clear_formula', 'Clear one cell formula.', { cell }, ['cell']),
  capability('xlsx.style.patch', 'Patch a bounded range using the native v1 style vocabulary.', { range, style }, ['range', 'style']),
  capability('xlsx.row.set_height', 'Set one zero-based row height in points.', { row: { type: 'integer', minimum: 0, maximum: 1_048_575 }, heightPoints: { type: 'number', minimum: 0, maximum: 409.5 } }, ['row', 'heightPoints']),
  capability('xlsx.column.set_width', 'Set one zero-based column width in Excel character units.', { column: { type: 'integer', minimum: 0, maximum: 16_383 }, width: { type: 'number', minimum: 0, maximum: 255 } }, ['column', 'width']),
  capability('xlsx.range.merge', 'Merge one rectangular range when native source authority permits it.', { range }, ['range']),
  capability('xlsx.range.unmerge', 'Unmerge one rectangular range when native source authority permits it.', { range }, ['range']),
]

function identity(artifact: XlsxAgentArtifact, workbook: XlsxNativeWorkbook): AgentArtifactIdentity {
  return { artifactId: artifact.artifactId, format: 'xlsx', mediaType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', revision: workbook.source.package_sha256, fingerprint: workbook.source.package_sha256 }
}
async function snapshot(artifact: XlsxAgentArtifact, signal?: AbortSignal): Promise<XlsxAgentSnapshot> {
  const current = await artifact.snapshot({ signal }); const decoded = current.workbook.version === 2 ? validateNativeWorkbookV2(current.workbook) : validateNativeWorkbookV1(current.workbook)
  if (!decoded.ok) throw new Error(`invalid native XLSX snapshot: ${decoded.issues[0]?.path ?? '/'} ${decoded.issues[0]?.message ?? ''}`.trim())
  return { workbook: structuredClone(decoded.value) }
}
function exact(input: JsonObject, allowed: string[]): void { for (const key of Object.keys(input)) if (!allowed.includes(key)) throw new Error(`unknown input field ${JSON.stringify(key)}`) }
function requiredString(input: JsonObject, key: string): string { const value = input[key]; if (typeof value !== 'string' || !value) throw new Error(`${key} must be a non-empty string`); return value }
function requiredNumber(input: JsonObject, key: string): number { const value = input[key]; if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error(`${key} must be a finite number`); return value }
function ref(input: JsonObject, key: 'cell'): { row: number; column: number }
function ref(input: JsonObject, key: 'range'): { row: number; column: number; end_row: number; end_column: number }
function ref(input: JsonObject, key: 'cell' | 'range') {
  const value = input[key]; if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${key} must be an object`)
  const record = value as JsonObject; exact(record, key === 'cell' ? ['row', 'column'] : ['row', 'column', 'endRow', 'endColumn'])
  return key === 'cell' ? { row: requiredNumber(record, 'row'), column: requiredNumber(record, 'column') } : { row: requiredNumber(record, 'row'), column: requiredNumber(record, 'column'), end_row: requiredNumber(record, 'endRow'), end_column: requiredNumber(record, 'endColumn') }
}
function mutation(operation: AgentOperation): SupportedWorkbookMutation {
  const input = operation.input; const sheet_id = requiredString(input, 'sheetId'); const base = { operation_id: operation.operationId, sheet_id }
  switch (operation.name) {
    case 'xlsx.cell.set_value': { exact(input, ['sheetId', 'cell', 'value']); const value = input.value; if (!['string', 'number', 'boolean'].includes(typeof value)) throw new Error('value must be a string, number, or boolean'); return { ...base, kind: 'cell.set_value', cell: ref(input, 'cell'), value: value as string | number | boolean } }
    case 'xlsx.cell.clear_value': exact(input, ['sheetId', 'cell']); return { ...base, kind: 'cell.clear_value', cell: ref(input, 'cell') }
    case 'xlsx.cell.set_formula': exact(input, ['sheetId', 'cell', 'formula']); return { ...base, kind: 'cell.set_formula', cell: ref(input, 'cell'), formula: requiredString(input, 'formula') }
    case 'xlsx.cell.clear_formula': exact(input, ['sheetId', 'cell']); return { ...base, kind: 'cell.clear_formula', cell: ref(input, 'cell') }
    case 'xlsx.style.patch': { exact(input, ['sheetId', 'range', 'style']); if (!input.style || typeof input.style !== 'object' || Array.isArray(input.style)) throw new Error('style must be an object'); if (Object.values(input.style).some((value) => value === null)) throw new Error('agent style patches do not accept null because direct-property clearing cannot be verified from the effective-style projection'); return { ...base, kind: 'style.patch', range: ref(input, 'range'), style: structuredClone(input.style) as StyleDelta } }
    case 'xlsx.row.set_height': exact(input, ['sheetId', 'row', 'heightPoints']); return { ...base, kind: 'row.set_height', row: requiredNumber(input, 'row'), height_points: requiredNumber(input, 'heightPoints') }
    case 'xlsx.column.set_width': exact(input, ['sheetId', 'column', 'width']); return { ...base, kind: 'column.set_width', column: requiredNumber(input, 'column'), width: requiredNumber(input, 'width') }
    case 'xlsx.range.merge': exact(input, ['sheetId', 'range']); return { ...base, kind: 'range.merge', range: ref(input, 'range') }
    case 'xlsx.range.unmerge': exact(input, ['sheetId', 'range']); return { ...base, kind: 'range.unmerge', range: ref(input, 'range') }
    default: throw new Error(`unsupported XLSX operation ${JSON.stringify(operation.name)}`)
  }
}
function batch(operations: readonly AgentOperation[], revision: string, id = 'agent-preview'): WorkbookMutationBatchV1 {
  const decoded = decodeWorkbookMutationBatch({ protocol: WORKBOOK_MUTATION_PROTOCOL, version: WORKBOOK_MUTATION_VERSION, batch_id: id, expected_revision: revision, operations: operations.map(mutation) })
  if (!decoded.ok) throw new Error(decoded.issues.map((issue) => `${issue.path || '/'}: ${issue.message}`).join('; ')); return decoded.value
}
function semanticIssues(workbook: XlsxNativeWorkbook, operations: readonly SupportedWorkbookMutation[]): AgentIssue[] {
  const issues: AgentIssue[] = []
  for (const operation of operations) {
    const sheet = workbook.sheets.find((candidate) => candidate.id === operation.sheet_id)
    if (!sheet) { issues.push({ severity: 'refusal', code: 'STALE_TARGET', operationId: operation.operation_id, path: '/input/sheetId', message: `sheet ${JSON.stringify(operation.sheet_id)} does not exist` }); continue }
    if (!sheet.editable) { issues.push({ severity: 'refusal', code: sheet.refusal_code ?? 'NATIVE_READ_ONLY', operationId: operation.operation_id, message: `sheet ${JSON.stringify(sheet.name)} is mutation-refused` }); continue }
    const area = 'cell' in operation ? { row: operation.cell.row, column: operation.cell.column, end_row: operation.cell.row, end_column: operation.cell.column } : 'range' in operation ? operation.range : undefined
    if (area && sheet.cells.some((cell) => !cell.editable && cell.row >= area.row && cell.row <= area.end_row && cell.column >= area.column && cell.column <= area.end_column)) issues.push({ severity: 'refusal', code: 'NATIVE_READ_ONLY', operationId: operation.operation_id, message: 'operation intersects a modeled mutation-refused cell' })
  }
  return issues
}
function prepared(workbook: XlsxNativeWorkbook, operations: readonly AgentOperation[], revision: string, id?: string): { value?: WorkbookMutationBatchV1; issues: AgentIssue[] } {
  try { const value = batch(operations, revision, id); return { value, issues: semanticIssues(workbook, value.operations) } } catch (error) { return { issues: [{ ...errorIssue(error, '/operations'), severity: 'refusal' }] } }
}
function cellValue(cell: XlsxNativeWorkbook['sheets'][number]['cells'][number] | undefined): unknown { if (!cell?.value) return undefined; if (cell.value.kind === 'string') return cell.value.text ?? ''; if (cell.value.kind === 'boolean') return cell.value.lexical === '1' || cell.value.lexical === 'true'; if (cell.value.kind === 'number') return Number(cell.value.lexical); return cell.value.lexical ?? cell.value.text }
function operationPassed(workbook: XlsxNativeWorkbook, operation: SupportedWorkbookMutation): boolean {
  const sheet = workbook.sheets.find((item) => item.id === operation.sheet_id); if (!sheet) return false
  const target = 'cell' in operation ? sheet.cells.find((item) => item.row === operation.cell.row && item.column === operation.cell.column) : undefined
  switch (operation.kind) {
    case 'cell.set_value': return target !== undefined && Object.is(cellValue(target), operation.value)
    case 'cell.clear_value': return target === undefined || target.value === undefined
    case 'cell.set_formula': return target?.formula?.text === operation.formula
    case 'cell.clear_formula': return target === undefined || target.formula === undefined
    case 'row.set_height': return sheet.rows.some((row) => row.row === operation.row && row.height_points === operation.height_points)
    case 'column.set_width': return sheet.columns.some((column) => operation.column >= column.column && operation.column <= column.end_column && column.width === operation.width)
    case 'range.merge': return sheet.merged_ranges.some((merged) => merged.row === operation.range.row && merged.column === operation.range.column && merged.end_row === operation.range.end_row && merged.end_column === operation.range.end_column)
    case 'range.unmerge': return !sheet.merged_ranges.some((merged) => merged.row === operation.range.row && merged.column === operation.range.column && merged.end_row === operation.range.end_row && merged.end_column === operation.range.end_column)
    case 'style.patch': {
      const cells = sheet.cells.filter((cell) => cell.row >= operation.range.row && cell.row <= operation.range.end_row && cell.column >= operation.range.column && cell.column <= operation.range.end_column)
      const expectedCount = (operation.range.end_row - operation.range.row + 1) * (operation.range.end_column - operation.range.column + 1)
      return cells.length === expectedCount && cells.every((cell) => { const style = workbook.styles.find((item) => item.id === cell.style_id)?.effective; return !!style && Object.entries(operation.style).every(([key, value]) => style[key as keyof typeof style] === value) })
    }
  }
}

export function createXlsxAgentAdapter(): AgentArtifactAdapter<XlsxAgentArtifact> {
  return {
    id: 'injoffice.xlsx.native-v1', format: 'xlsx',
    supports(value): value is XlsxAgentArtifact { return typeof value === 'object' && value !== null && (value as Partial<XlsxAgentArtifact>).kind === 'injoffice.xlsx' && typeof (value as Partial<XlsxAgentArtifact>).artifactId === 'string' && typeof (value as Partial<XlsxAgentArtifact>).snapshot === 'function' && typeof (value as Partial<XlsxAgentArtifact>).apply === 'function' },
    async identity({ artifact, signal }) { const current = await snapshot(artifact, signal); return identity(artifact, current.workbook) },
    async capabilities() { return structuredClone(CAPABILITIES) },
    async inspect({ artifact, cursor, maxItems, maxBytes, signal }) { const { workbook } = await snapshot(artifact, signal); return boundedObject({ documentId: workbook.document_id, projectionRevision: workbook.revision, sheets: toJson(workbook.sheets.map((s) => ({ id: s.id, name: s.name, state: s.state, editable: s.editable, cellCount: s.cells.length }))), capabilities: toJson(workbook.capabilities) }, 'unsupported', workbook.unsupported, cursorOffset(cursor), maxItems, maxBytes) },
    async read({ artifact, query, cursor, maxItems, maxBytes, signal }) { const { workbook } = await snapshot(artifact, signal); const sheetId = query?.sheetId; const sheetName = query?.sheetName; const selected = workbook.sheets.find((item) => typeof sheetId === 'string' ? item.id === sheetId : typeof sheetName === 'string' ? item.name === sheetName : true); if (!selected) throw new Error('requested XLSX sheet does not exist'); return boundedObject({ sheet: toJson({ id: selected.id, name: selected.name, editable: selected.editable }) }, 'cells', selected.cells, cursorOffset(cursor), maxItems, maxBytes) },
    async validate(request) { const current = await snapshot(request.artifact, request.signal); const issues: AgentIssue[] = []; try { assertFresh(identity(request.artifact, current.workbook), request.sourceIdentity) } catch (error) { issues.push({ ...errorIssue(error), severity: 'error' }) }; issues.push(...prepared(current.workbook, request.changeSet.operations, request.changeSet.baseRevision).issues); return issues },
    async preview(request) { const current = await snapshot(request.artifact, request.signal); assertFresh(identity(request.artifact, current.workbook), request.sourceIdentity); const result = prepared(current.workbook, request.changeSet.operations, request.changeSet.baseRevision); if (!result.value || result.issues.length) return { data: toJson({ mode: 'refused' }), issues: result.issues, evidence: [] }; if (!request.artifact.preview) return { data: toJson({ mode: 'validated-plan', operationCount: result.value.operations.length }), issues: [], evidence: [{ kind: 'xlsx.mutation-batch', data: toJson(result.value) }] }; const candidate = await request.artifact.preview(result.value, { signal: request.signal }); const checked = candidate.workbook.version === 2 ? validateNativeWorkbookV2(candidate.workbook) : validateNativeWorkbookV1(candidate.workbook); if (!checked.ok) throw new Error(`native XLSX preview returned invalid output: ${checked.issues[0]?.message ?? 'unknown error'}`); const passed = result.value.operations.every((operation) => operationPassed(checked.value, operation)); return { data: toJson({ mode: 'native-host', operationCount: result.value.operations.length, readbackPassed: passed, resultingRevision: checked.value.source.package_sha256 }), issues: passed ? [] : [{ severity: 'error', code: 'PREVIEW_MISMATCH', message: 'preview readback did not contain every requested XLSX mutation' }], evidence: [{ kind: 'xlsx.fresh-native-preview', data: { documentId: checked.value.document_id, sheets: checked.value.sheets.length } }] } },
    async diff(request) { const current = await snapshot(request.artifact, request.signal); const result = prepared(current.workbook, request.changeSet.operations, request.changeSet.baseRevision); return { data: { operationCount: result.value?.operations.length ?? 0, touchedSheets: result.value ? [...new Set(result.value.operations.map((op) => op.sheet_id))] : [], operations: result.value ? toJson(result.value.operations) : [] }, issues: result.issues, evidence: [] } },
    async commit(request) { const current = await snapshot(request.artifact, request.signal); assertFresh(identity(request.artifact, current.workbook), request.sourceIdentity); const result = prepared(current.workbook, request.changeSet.operations, request.expectedRevision, request.idempotencyKey); if (!result.value || result.issues.length) throw new Error(result.issues.map((issue) => issue.message).join('; ')); const applied = await request.artifact.apply(result.value, { signal: request.signal }); const checked = applied.snapshot.workbook.version === 2 ? validateNativeWorkbookV2(applied.snapshot.workbook) : validateNativeWorkbookV1(applied.snapshot.workbook); if (!checked.ok) throw new Error(`native XLSX host returned invalid output: ${checked.issues[0]?.message ?? 'unknown error'}`); return { artifact: request.artifact, identity: identity(request.artifact, checked.value), data: { batchId: result.value.batch_id, operationCount: result.value.operations.length, ...(applied.receipt === undefined ? {} : { hostReceipt: applied.receipt }) }, evidence: [{ kind: 'xlsx.native-commit', data: { operationIds: result.value.operations.map((op) => op.operation_id) } }], deduplicated: applied.deduplicated } },
    async verify(request) { if (request.stage === 'planned') { const issues = await this.validate(request); const passed = !issues.some((issue) => issue.severity !== 'warning'); return { verified: passed, checks: [{ name: 'native-xlsx-plan', passed, evidence: [{ kind: 'xlsx.validation', data: { issueCount: issues.length } }] }], issues } }; const current = await snapshot(request.artifact, request.signal); const native = prepared(current.workbook, request.changeSet.operations, request.changeSet.baseRevision).value; const checks = native?.operations.map((operation) => ({ name: `operation:${operation.operation_id}`, passed: operationPassed(current.workbook, operation), evidence: [{ kind: 'xlsx.readback', data: { operationId: operation.operation_id, kind: operation.kind } }] })) ?? []; const issues = checks.filter((check) => !check.passed).map((check) => ({ severity: 'error' as const, code: 'VERIFY_FAILED', message: `${check.name} was not observed in the fresh native projection` })); return { verified: checks.length > 0 && checks.every((check) => check.passed), checks, issues } },
  }
}
