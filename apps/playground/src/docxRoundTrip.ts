import type {
  NativeDocxDocumentV1,
  NativeDocxParagraphV1,
  NativeDocxStoryV1,
} from '../../../packages/docs/src/nativeContract'
import {
  OFFICE_MUTATION_PROTOCOL,
  OFFICE_MUTATION_VERSION,
  type NativeDocxOfficeMutationEnvelopeV1,
} from '../../../packages/docs/src/nativeTransactionAdapterV1'

export type EditableDocxRun = {
  key: string
  label: string
  paragraphId: string
  runId: string
  partName: string
  text: string
  expectedXmlSHA256: string
}

export type DocxMutationEvidence = {
  operation: 'text.replace'
  target_kind: 'run'
  target_id: string
  source_part: string
  expected_xml_sha256: string
  expected_revision: string
  replacement_utf16_units: number
  replacement_preview: string
}

export type DocxRoundTripProof = {
  target: string
  before: string
  after: string
  documentIdentity: 'stable' | 'reissued by exact-byte readback'
  artifactIdentity: 'stable' | 'browser memory' | 'request-local'
  previousRevision: string
  revision: string
  preservedParts: number
}

function targetKey(partName: string, runId: string): string {
  return `${encodeURIComponent(partName)}:${runId}`
}

function visitStory(story: NativeDocxStoryV1, storyLabel: string, targets: EditableDocxRun[]): void {
  const visit = (paragraph: NativeDocxParagraphV1, location: string) => {
    if (paragraph.edit_policy.mode !== 'read-write' || !paragraph.edit_policy.allowed_operations.includes('text.replace')) return
    paragraph.runs.forEach((run, runIndex) => {
      if (run.kind !== 'text' || run.text === undefined || run.properties?.hidden) return
      targets.push({
        key: targetKey(run.anchor.part_name, run.id),
        label: `${storyLabel} · ${location} · run ${runIndex + 1}`,
        paragraphId: paragraph.id,
        runId: run.id,
        partName: run.anchor.part_name,
        text: run.text,
        expectedXmlSHA256: run.anchor.xml_sha256,
      })
    })
  }

  story.blocks.forEach((block, blockIndex) => {
    if (block.paragraph) {
      visit(block.paragraph, `paragraph ${blockIndex + 1}`)
      return
    }
    block.table?.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
      cell.paragraphs.forEach((paragraph, paragraphIndex) => {
        visit(paragraph, `table ${blockIndex + 1}, row ${rowIndex + 1}, cell ${cellIndex + 1}, paragraph ${paragraphIndex + 1}`)
      })
    }))
  })
}

export function editableDocxRuns(document: NativeDocxDocumentV1): EditableDocxRun[] {
  const targets: EditableDocxRun[] = []
  visitStory(document.body, 'Body', targets)
  document.headers.forEach((story, index) => visitStory(story, `Header ${index + 1}`, targets))
  document.footers.forEach((story, index) => visitStory(story, `Footer ${index + 1}`, targets))
  document.notes.forEach((story, index) => visitStory(story, `Note ${index + 1}`, targets))
  document.comment_stories.forEach((story, index) => visitStory(story, `Comment ${index + 1}`, targets))
  return targets
}

export function buildDocxRunMutation(
  document: NativeDocxDocumentV1,
  target: EditableDocxRun,
  text: string,
  mutationId: string,
): NativeDocxOfficeMutationEnvelopeV1 {
  return {
    protocol: OFFICE_MUTATION_PROTOCOL,
    version: OFFICE_MUTATION_VERSION,
    format: 'docx',
    mutation_id: mutationId,
    expected_revision: document.source.package_sha256,
    payload: {
      mutations: [{
        target_kind: 'run',
        target_id: target.runId,
        expected_xml_sha256: target.expectedXmlSHA256,
        text,
      }],
    },
  }
}

export function buildDocxMutationEvidence(document: NativeDocxDocumentV1, target: EditableDocxRun, text: string): DocxMutationEvidence {
  const previewLimit = 120
  return {
    operation: 'text.replace',
    target_kind: 'run',
    target_id: target.runId,
    source_part: target.partName,
    expected_xml_sha256: target.expectedXmlSHA256,
    expected_revision: document.source.package_sha256,
    replacement_utf16_units: text.length,
    replacement_preview: text.length > previewLimit ? `${text.slice(0, previewLimit)}…` : text,
  }
}

export function findEditableDocxRun(document: NativeDocxDocumentV1, target: Pick<EditableDocxRun, 'runId' | 'partName'>): EditableDocxRun | undefined {
  return editableDocxRuns(document).find((candidate) => candidate.runId === target.runId && candidate.partName === target.partName)
}

function passthroughInventory(document: NativeDocxDocumentV1): string[] {
  return document.passthrough_parts
    .map((part) => `${part.part_name}\0${part.sha256}\0${part.byte_length}\0${part.content_type}\0${part.policy}`)
    .sort()
}

export function verifyDocxRoundTrip(
  before: NativeDocxDocumentV1,
  after: NativeDocxDocumentV1,
  target: EditableDocxRun,
  replacement: string,
  responseRevision?: string,
  responsePackageSHA256?: string,
  beforeArtifactId = '',
  responseArtifactId?: string,
  browserLocal = false,
): DocxRoundTripProof {
  const nextTarget = findEditableDocxRun(after, target)
  if (!nextTarget || nextTarget.text !== replacement) throw new Error(`Readback mismatch at ${target.label}.`)
  if (after.source.main_part !== before.source.main_part) throw new Error('Readback changed the native main-part identity.')
  if (after.revision === before.revision) throw new Error('Readback revision did not change after the text mutation.')
  if (after.source.package_sha256 === before.source.package_sha256) throw new Error('Readback package digest did not change after the text mutation.')
  if (responseRevision && responseRevision !== after.revision) throw new Error('Readback revision does not match the saved package.')
  if (responsePackageSHA256 && responsePackageSHA256 !== after.source.package_sha256) throw new Error('Readback digest does not match the saved package.')
  if (beforeArtifactId && responseArtifactId !== beforeArtifactId) throw new Error('Mutation response changed the server artifact identity.')

  const preservedBefore = passthroughInventory(before)
  const preservedAfter = passthroughInventory(after)
  if (JSON.stringify(preservedAfter) !== JSON.stringify(preservedBefore)) throw new Error('Readback changed the preserve-verbatim part inventory.')

  return {
    target: target.label,
    before: target.text,
    after: nextTarget.text,
    documentIdentity: after.document_id === before.document_id ? 'stable' : 'reissued by exact-byte readback',
    artifactIdentity: browserLocal ? 'browser memory' : beforeArtifactId ? 'stable' : 'request-local',
    previousRevision: before.source.package_sha256,
    revision: after.source.package_sha256,
    preservedParts: preservedAfter.length,
  }
}

export function docxDownloadName(sourceName: string): string {
  const base = sourceName.replace(/\.docx$/i, '').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '')
  return `${base || 'document'}-injoffice.docx`
}
