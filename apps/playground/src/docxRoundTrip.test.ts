import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { decodeNativeDocx } from './docsNativePreview'
import {
  buildDocxMutationEvidence,
  buildDocxRunMutation,
  docxDownloadName,
  editableDocxRuns,
  findEditableDocxRun,
  verifyDocxRoundTrip,
} from './docxRoundTrip'

const record = JSON.parse(readFileSync(
  new URL('../../../go/officecompat/corpus/generated/expected/docx-transitional-common.json', import.meta.url),
  'utf8',
)) as { native: unknown }
const document = decodeNativeDocx(record.native)

describe('native DOCX round trip', () => {
  it('selects guarded visible text runs from body paragraphs and table cells', () => {
    const targets = editableDocxRuns(document)
    expect(targets.length).toBeGreaterThan(1)
    expect(targets.some((target) => target.label.startsWith('Body · paragraph '))).toBe(true)
    expect(targets.some((target) => target.label.startsWith('Body · table '))).toBe(true)
    expect(targets.every((target) => target.text.length > 0 && target.expectedXmlSHA256.startsWith('sha256:'))).toBe(true)
  })

  it('excludes hidden text and text owned by a read-only paragraph', () => {
    const changed = structuredClone(document)
    const paragraphs = changed.body.blocks.flatMap((block) => block.paragraph
      ? [block.paragraph]
      : block.table?.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) ?? [])
    const first = paragraphs.find((paragraph) => paragraph.runs.some((run) => run.kind === 'text'))!
    const firstRun = first.runs.find((run) => run.kind === 'text')!
    firstRun.properties = { ...firstRun.properties, hidden: true }
    const second = paragraphs.find((paragraph) => paragraph !== first && paragraph.runs.some((run) => run.kind === 'text'))!
    second.edit_policy = { mode: 'read-only', allowed_operations: [], refusal: { code: 'READ_ONLY', message: 'test', preservation: 'refuse-mutation' } }

    const targets = editableDocxRuns(changed)
    expect(targets.some((target) => target.runId === firstRun.id)).toBe(false)
    expect(targets.some((target) => target.paragraphId === second.id)).toBe(false)
  })

  it('builds one exact package- and anchor-bound run mutation', () => {
    const target = editableDocxRuns(document)[0]!
    expect(buildDocxRunMutation(document, target, 'After', 'save-1')).toEqual({
      protocol: 'injoffice.office.mutations',
      version: 1,
      format: 'docx',
      mutation_id: 'save-1',
      expected_revision: document.source.package_sha256,
      payload: { mutations: [{
        target_kind: 'run',
        target_id: target.runId,
        expected_xml_sha256: target.expectedXmlSHA256,
        text: 'After',
      }] },
    })
    expect(findEditableDocxRun(document, target)).toEqual(target)
    expect(buildDocxMutationEvidence(document, target, 'After')).toEqual({
      operation: 'text.replace',
      target_kind: 'run',
      target_id: target.runId,
      source_part: target.partName,
      expected_xml_sha256: target.expectedXmlSHA256,
      expected_revision: document.source.package_sha256,
      replacement_utf16_units: 5,
      replacement_preview: 'After',
    })
  })

  it('bounds evidence and verifies exact-byte readback before download', () => {
    const target = editableDocxRuns(document)[0]!
    expect(buildDocxMutationEvidence(document, target, 'x'.repeat(200)).replacement_preview).toHaveLength(121)
    const after = structuredClone(document)
    const paragraph = after.body.blocks.flatMap((block) => block.paragraph
      ? [block.paragraph]
      : block.table?.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs)) ?? [])
      .find((candidate) => candidate.id === target.paragraphId)!
    paragraph.runs.find((run) => run.id === target.runId)!.text = 'After'
    after.document_id = 'document:reissued'
    after.revision = 'rev:22222222222222222222222222222222'
    after.source.package_sha256 = `sha256:${'2'.repeat(64)}`

    expect(verifyDocxRoundTrip(document, after, target, 'After', after.revision, after.source.package_sha256, 'artifact-1', 'artifact-1')).toMatchObject({
      after: 'After',
      documentIdentity: 'reissued by exact-byte readback',
      artifactIdentity: 'stable',
      preservedParts: document.passthrough_parts.length,
    })
    after.passthrough_parts[0]!.sha256 = `sha256:${'f'.repeat(64)}`
    expect(() => verifyDocxRoundTrip(document, after, target, 'After')).toThrow(/preserve-verbatim/)
  })

  it('creates a safe output filename', () => {
    expect(docxDownloadName('Quarterly plan.docx')).toBe('Quarterly-plan-injoffice.docx')
    expect(docxDownloadName('***.docx')).toBe('document-injoffice.docx')
  })
})
