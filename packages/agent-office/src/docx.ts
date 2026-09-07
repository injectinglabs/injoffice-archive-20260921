import {
  OFFICE_MUTATION_PROTOCOL, OFFICE_MUTATION_VERSION, decodeNativeDocxDocument,
  type NativeDocxDocumentV1, type NativeDocxOfficeMutationEnvelopeV1, type NativeDocxParagraphV1, type NativeDocxRunV1,
} from '@injoffice/docs/native-docx'
import type { AgentArtifactAdapter, AgentArtifactIdentity, AgentCapability, AgentIssue, AgentOperation, JsonValue } from '@injoffice/agent-tools'
import { assertFresh, boundedObject, cursorOffset, errorIssue, toJson } from './hash.js'

export interface DocxAgentSnapshot { document: NativeDocxDocumentV1 }
export interface DocxAgentArtifact {
  readonly kind: 'injoffice.docx'; readonly artifactId: string
  snapshot(options?: { signal?: AbortSignal }): Promise<DocxAgentSnapshot>
  /** Must atomically enforce expected_revision and durably deduplicate mutation_id. */
  apply(envelope: NativeDocxOfficeMutationEnvelopeV1, options?: { signal?: AbortSignal }): Promise<{ snapshot: DocxAgentSnapshot; receipt?: JsonValue; deduplicated?: boolean }>
  preview?(envelope: NativeDocxOfficeMutationEnvelopeV1, options?: { signal?: AbortSignal }): Promise<DocxAgentSnapshot>
}

const CAPABILITY: AgentCapability = {
  name: 'docx.text.replace',
  description: 'Replace the complete text of one stable, source-anchored native DOCX run or eligible single-run paragraph.',
  destructive: true, requiresConfirmation: true,
  inputSchema: {
    type: 'object', additionalProperties: false,
    properties: { targetKind: { type: 'string', enum: ['run', 'paragraph'] }, targetId: { type: 'string', minLength: 1, maxLength: 256 }, text: { type: 'string', maxLength: 1_048_576 }, expectedText: { type: 'string', maxLength: 1_048_576 } },
    required: ['targetKind', 'targetId', 'text'],
  },
  metadata: { execution: 'host', wholeFileWrite: true, sourceAnchored: true },
}

function identity(artifact: DocxAgentArtifact, document: NativeDocxDocumentV1): AgentArtifactIdentity { return { artifactId: artifact.artifactId, format: 'docx', mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', revision: document.source.package_sha256, fingerprint: document.source.package_sha256 } }
async function snapshot(artifact: DocxAgentArtifact, signal?: AbortSignal): Promise<DocxAgentSnapshot> { const current = await artifact.snapshot({ signal }); const decoded = decodeNativeDocxDocument(current.document); if (!decoded.ok) throw new Error(`invalid native DOCX snapshot: ${decoded.issues[0]?.path ?? '/'} ${decoded.issues[0]?.message ?? ''}`.trim()); return { document: structuredClone(decoded.value) } }
function stories(document: NativeDocxDocumentV1) { return [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories] }
function target(document: NativeDocxDocumentV1, kind: 'paragraph' | 'run', id: string): { paragraph: NativeDocxParagraphV1; run?: NativeDocxRunV1; text: string; hash: string } | undefined {
  for (const story of stories(document)) for (const block of story.blocks) {
    const paragraph = block.paragraph; if (!paragraph) continue
    if (kind === 'paragraph' && paragraph.id === id && paragraph.runs.length === 1 && paragraph.runs[0]?.kind === 'text') return { paragraph, run: paragraph.runs[0], text: paragraph.runs[0].text ?? '', hash: paragraph.anchor.xml_sha256 }
    if (kind === 'run') { const run = paragraph.runs.find((candidate) => candidate.id === id && candidate.kind === 'text'); if (run) return { paragraph, run, text: run.text ?? '', hash: run.anchor.xml_sha256 } }
  }
  return undefined
}
function xmlTextValid(value: string): boolean { if (value.length > 1_048_576) return false; for (let index = 0; index < value.length;) { const code = value.codePointAt(index)!; if (!(code === 9 || code === 10 || code === 13 || code >= 0x20 && code <= 0xd7ff || code >= 0xe000 && code <= 0xfffd || code >= 0x10000 && code <= 0x10ffff) || code === 0xfffe || code === 0xffff) return false; index += code > 0xffff ? 2 : 1 } return true }
function envelope(document: NativeDocxDocumentV1, operations: readonly AgentOperation[], mutationId: string): { value?: NativeDocxOfficeMutationEnvelopeV1; issues: AgentIssue[] } {
  const issues: AgentIssue[] = []; const mutations: NativeDocxOfficeMutationEnvelopeV1['payload']['mutations'] = []; const targets = new Map<string, string>()
  for (const operation of operations) {
    if (operation.name !== CAPABILITY.name) { issues.push({ severity: 'refusal', code: 'UNSUPPORTED_OPERATION', operationId: operation.operationId, message: `unsupported DOCX operation ${operation.name}` }); continue }
    const allowed = ['targetKind', 'targetId', 'text', 'expectedText']; const extra = Object.keys(operation.input).find((key) => !allowed.includes(key)); if (extra) { issues.push({ severity: 'refusal', code: 'UNKNOWN_FIELD', operationId: operation.operationId, path: `/input/${extra}`, message: `unknown input field ${extra}` }); continue }
    const { targetKind, targetId, text, expectedText } = operation.input
    if ((targetKind !== 'run' && targetKind !== 'paragraph') || typeof targetId !== 'string' || !targetId || typeof text !== 'string' || !xmlTextValid(text) || expectedText !== undefined && typeof expectedText !== 'string') { issues.push({ severity: 'refusal', code: 'INVALID_INPUT', operationId: operation.operationId, message: 'targetKind, targetId, and XML-safe bounded text are required' }); continue }
    const found = target(document, targetKind, targetId)
    if (!found) { issues.push({ severity: 'refusal', code: 'UNSUPPORTED_TARGET', operationId: operation.operationId, message: `native ${targetKind} ${JSON.stringify(targetId)} is absent or not safely text-replaceable` }); continue }
    if (found.paragraph.edit_policy.mode !== 'read-write' || !found.paragraph.edit_policy.allowed_operations.includes('text.replace')) { issues.push({ severity: 'refusal', code: found.paragraph.edit_policy.refusal?.code ?? 'NATIVE_READ_ONLY', operationId: operation.operationId, message: found.paragraph.edit_policy.refusal?.message ?? 'target paragraph is mutation-refused' }); continue }
    if (expectedText !== undefined && expectedText !== found.text) { issues.push({ severity: 'error', code: 'STALE_TARGET', operationId: operation.operationId, message: 'expectedText does not match the exact native target' }); continue }
    if (text === found.text) { issues.push({ severity: 'refusal', code: 'SEMANTIC_NO_OP', operationId: operation.operationId, message: 'replacement text is unchanged' }); continue }
    const paragraphKey = found.paragraph.id; if (targets.has(`${targetKind}:${targetId}`) || targets.has(`paragraph:${paragraphKey}`) || targetKind === 'paragraph' && [...targets.values()].includes(paragraphKey)) { issues.push({ severity: 'refusal', code: 'AMBIGUOUS_EDIT', operationId: operation.operationId, message: 'a DOCX target or its owning paragraph is targeted more than once' }); continue }
    targets.set(`${targetKind}:${targetId}`, paragraphKey)
    mutations.push({ target_kind: targetKind, target_id: targetId, expected_xml_sha256: found.hash, text })
  }
  if (issues.length) return { issues }
  return { value: { protocol: OFFICE_MUTATION_PROTOCOL, version: OFFICE_MUTATION_VERSION, format: 'docx', mutation_id: mutationId, expected_revision: document.source.package_sha256, payload: { mutations } }, issues }
}
function readback(document: NativeDocxDocumentV1, operation: AgentOperation): boolean { const kind = operation.input.targetKind; const id = operation.input.targetId; const text = operation.input.text; return (kind === 'run' || kind === 'paragraph') && typeof id === 'string' && typeof text === 'string' && target(document, kind, id)?.text === text }

export function createDocxAgentAdapter(): AgentArtifactAdapter<DocxAgentArtifact> {
  return {
    id: 'injoffice.docx.native-text-v1', format: 'docx',
    supports(value): value is DocxAgentArtifact { return typeof value === 'object' && value !== null && (value as Partial<DocxAgentArtifact>).kind === 'injoffice.docx' && typeof (value as Partial<DocxAgentArtifact>).artifactId === 'string' && typeof (value as Partial<DocxAgentArtifact>).snapshot === 'function' && typeof (value as Partial<DocxAgentArtifact>).apply === 'function' },
    async identity({ artifact, signal }) { const current = await snapshot(artifact, signal); return identity(artifact, current.document) },
    async capabilities() { return [structuredClone(CAPABILITY)] },
    async inspect({ artifact, cursor, maxItems, maxBytes, signal }) { const { document } = await snapshot(artifact, signal); return boundedObject({ documentId: document.document_id, projectionRevision: document.revision, stories: toJson(stories(document).map((story) => ({ id: story.id, kind: story.kind, blockCount: story.blocks.length }))), capabilities: toJson(document.capabilities) }, 'unsupported', document.unsupported, cursorOffset(cursor), maxItems, maxBytes) },
    async read({ artifact, query, cursor, maxItems, maxBytes, signal }) { const { document } = await snapshot(artifact, signal); const storyId = query?.storyId; const story = typeof storyId === 'string' ? stories(document).find((item) => item.id === storyId) : document.body; if (!story) throw new Error('requested DOCX story does not exist'); return boundedObject({ story: toJson({ id: story.id, kind: story.kind }) }, 'blocks', story.blocks, cursorOffset(cursor), maxItems, maxBytes) },
    async validate(request) { const current = await snapshot(request.artifact, request.signal); const issues: AgentIssue[] = []; try { assertFresh(identity(request.artifact, current.document), request.sourceIdentity) } catch (error) { issues.push({ ...errorIssue(error), severity: 'error' }) }; issues.push(...envelope(current.document, request.changeSet.operations, request.changeSet.changeSetId).issues); return issues },
    async preview(request) { const current = await snapshot(request.artifact, request.signal); assertFresh(identity(request.artifact, current.document), request.sourceIdentity); const result = envelope(current.document, request.changeSet.operations, request.changeSet.changeSetId); if (!result.value) return { data: toJson({ mode: 'refused' }), issues: result.issues, evidence: [] }; if (!request.artifact.preview) return { data: toJson({ mode: 'validated-envelope', mutationCount: result.value.payload.mutations.length }), issues: [], evidence: [{ kind: 'docx.mutation-envelope', data: toJson(result.value) }] }; const candidate = await request.artifact.preview(result.value, { signal: request.signal }); const decoded = decodeNativeDocxDocument(candidate.document); if (!decoded.ok) throw new Error(`native DOCX preview returned invalid output: ${decoded.issues[0]?.message ?? 'unknown error'}`); const passed = request.changeSet.operations.every((operation) => readback(decoded.value, operation)); return { data: toJson({ mode: 'native-host', mutationCount: result.value.payload.mutations.length, readbackPassed: passed, resultingRevision: decoded.value.source.package_sha256 }), issues: passed ? [] : [{ severity: 'error', code: 'PREVIEW_MISMATCH', message: 'preview readback did not contain every requested replacement' }], evidence: [{ kind: 'docx.fresh-native-preview', data: { documentId: decoded.value.document_id } }] } },
    async diff(request) { const current = await snapshot(request.artifact, request.signal); const result = envelope(current.document, request.changeSet.operations, request.changeSet.changeSetId); return { data: { mutations: result.value ? toJson(result.value.payload.mutations.map(({ target_kind, target_id, text }) => ({ targetKind: target_kind, targetId: target_id, replacementText: text }))) : [] }, issues: result.issues, evidence: [] } },
    async commit(request) { const current = await snapshot(request.artifact, request.signal); assertFresh(identity(request.artifact, current.document), request.sourceIdentity); const result = envelope(current.document, request.changeSet.operations, request.idempotencyKey); if (!result.value) throw new Error(result.issues.map((issue) => issue.message).join('; ')); const applied = await request.artifact.apply(result.value, { signal: request.signal }); const decoded = decodeNativeDocxDocument(applied.snapshot.document); if (!decoded.ok) throw new Error(`native DOCX host returned invalid output: ${decoded.issues[0]?.message ?? 'unknown error'}`); return { artifact: request.artifact, identity: identity(request.artifact, decoded.value), data: { mutationId: result.value.mutation_id, mutationCount: result.value.payload.mutations.length, ...(applied.receipt === undefined ? {} : { hostReceipt: applied.receipt }) }, evidence: [{ kind: 'docx.native-commit', data: { targets: result.value.payload.mutations.map((item) => item.target_id) } }], deduplicated: applied.deduplicated } },
    async verify(request) { if (request.stage === 'planned') { const issues = await this.validate(request); const passed = !issues.some((issue) => issue.severity !== 'warning'); return { verified: passed, checks: [{ name: 'native-docx-plan', passed }], issues } }; const current = await snapshot(request.artifact, request.signal); const checks = request.changeSet.operations.map((operation) => ({ name: `operation:${operation.operationId}`, passed: readback(current.document, operation), evidence: [{ kind: 'docx.readback', data: { operationId: operation.operationId, targetId: operation.input.targetId ?? null } }] })); const issues = checks.filter((check) => !check.passed).map((check) => ({ severity: 'error' as const, code: 'VERIFY_FAILED', message: `${check.name} was not observed in the fresh native projection` })); return { verified: checks.length > 0 && checks.every((check) => check.passed), checks, issues } },
  }
}
