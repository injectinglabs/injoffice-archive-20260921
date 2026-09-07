import { readFileSync } from 'node:fs'
import { createAgentSession } from '@injoffice/agent-tools'
import type { WorkbookMutationBatchV1 } from '@injoffice/sheets'
import { describe, expect, it } from 'vitest'
import { createXlsxAgentAdapter, type XlsxAgentArtifact, type XlsxNativeWorkbook } from './xlsx.js'

const fixture = JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-get-corpus/pass-agent-dejavu.workbook.json', import.meta.url), 'utf8')) as XlsxNativeWorkbook
function hostArtifact(): XlsxAgentArtifact & { batches: WorkbookMutationBatchV1[] } {
  let workbook = structuredClone(fixture); const batches: WorkbookMutationBatchV1[] = []
  const apply = (batch: WorkbookMutationBatchV1) => {
    const next = structuredClone(workbook); const operation = batch.operations[0]!; if (operation.kind !== 'cell.set_value') throw new Error('test host only supports cell.set_value')
    const cell = next.sheets.find((sheet) => sheet.id === operation.sheet_id)!.cells.find((item) => item.row === operation.cell.row && item.column === operation.cell.column)!
    ;(cell as { value: unknown }).value = { kind: 'string', storage: 'inline', text: operation.value as string, rich: false }; (cell as { ooxml_type: string }).ooxml_type = 'inlineStr'; delete (cell as { formula?: unknown }).formula
    const digest = 'a'.repeat(64); const oldRevision = next.revision; const oldPackage = next.source.package_sha256
    const rebind = (value: unknown): void => { if (!value || typeof value !== 'object') return; for (const [key, child] of Object.entries(value)) { if (child === oldRevision) (value as Record<string, unknown>)[key] = `rev:${digest}`; else if (child === oldPackage) (value as Record<string, unknown>)[key] = `sha256:${digest}`; else rebind(child) } }; rebind(next)
    workbook = next; batches.push(structuredClone(batch)); return { snapshot: { workbook }, receipt: { saved: true } }
  }
  return { kind: 'injoffice.xlsx', artifactId: 'xlsx:1', batches, async snapshot() { return { workbook } }, async preview(batch) { const before = workbook; const result = apply(batch); workbook = before; batches.pop(); return result.snapshot }, async apply(batch) { return apply(batch) } }
}

describe('XLSX agent session', () => {
  it('runs the complete v2 native changeset lifecycle', async () => {
    const artifact = hostArtifact(); const session = await createAgentSession({ artifact, actor: { id: 'agent-1', kind: 'agent' }, adapter: createXlsxAgentAdapter(), confirmDestructive: () => true })
    expect((await session.capabilities()).capabilities.some(({ name }) => name === 'xlsx.cell.set_value')).toBe(true)
    expect((await session.inspect({ maxItems: 5, maxBytes: 20_000 })).identity.format).toBe('xlsx')
    expect((await session.read({ query: { sheetId: '1' }, maxItems: 2, maxBytes: 20_000 })).itemCount).toBe(2)
    const change = await session.plan([{ name: 'xlsx.cell.set_value', input: { sheetId: '1', cell: { row: 1, column: 0 }, value: 'Agent value' } }])
    expect((await change.validate()).valid).toBe(true); expect((await change.preview()).data).toMatchObject({ mode: 'native-host' }); expect((await change.diff()).data).toMatchObject({ operationCount: 1 })
    const committed = await change.commit({ idempotencyKey: 'xlsx-save-1', confirmation: { approved: true } })
    expect(artifact.batches[0]).toMatchObject({ batch_id: 'xlsx-save-1', operations: [{ operation_id: 'op-0001' }] }); expect(committed.verification.verified).toBe(true)
  })
  it('refuses malformed targets and prevents a non-advancing commit', async () => {
    const artifact = hostArtifact(); const session = await createAgentSession({ artifact, actor: { id: 'agent-1', kind: 'agent' }, adapter: createXlsxAgentAdapter(), confirmDestructive: () => true })
    const change = await session.plan([{ name: 'xlsx.cell.clear_value', input: { sheetId: 'missing', cell: { row: 0, column: 0 } } }]); const report = await change.validate()
    expect(report.valid).toBe(false); expect(report.issues.some(({ code }) => code === 'STALE_TARGET')).toBe(true); expect(artifact.batches).toHaveLength(0)
  })
  it('rejects unverifiable style clears and detects a lying preview host', async () => {
    const artifact = hostArtifact(); artifact.preview = async () => artifact.snapshot()
    const session = await createAgentSession({ artifact, actor: { id: 'agent-1', kind: 'agent' }, adapter: createXlsxAgentAdapter(), confirmDestructive: () => true })
    const clearStyle = await session.plan([{ name: 'xlsx.style.patch', input: { sheetId: '1', range: { row: 0, column: 0, endRow: 0, endColumn: 0 }, style: { bold: null } } }])
    expect((await clearStyle.validate()).valid).toBe(false)
    const value = await session.plan([{ name: 'xlsx.cell.set_value', input: { sheetId: '1', cell: { row: 1, column: 0 }, value: 'not-applied' } }])
    const preview = await value.preview(); expect(preview.data).toMatchObject({ readbackPassed: false }); expect(preview.issues).toContainEqual(expect.objectContaining({ code: 'PREVIEW_MISMATCH' }))
  })
})
