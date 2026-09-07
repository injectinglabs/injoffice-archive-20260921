import { applyPageOps, readInfo, type PageOpSpec, type PdfDocumentInfo } from '@injoffice/pdf'
import type { AgentArtifactAdapter, AgentArtifactIdentity, AgentCapability, AgentIssue, AgentOperation, JsonObject } from '@injoffice/agent-tools'
import { assertFresh, boundedObject, cursorOffset, errorIssue, fingerprintBytes, toJson } from './hash.js'

export interface PdfAgentArtifact { readonly kind: 'injoffice.pdf'; readonly artifactId: string; readonly bytes: Uint8Array }
const selector: JsonObject = { oneOf: [{ const: 'all' }, { type: 'array', minItems: 1, items: { type: 'integer', minimum: 1 }, uniqueItems: true }] }
const pageSize: JsonObject = {
  type: 'object',
  properties: { width: { type: 'number', exclusiveMinimum: 0 }, height: { type: 'number', exclusiveMinimum: 0 } },
  required: ['width', 'height'],
  additionalProperties: false,
}
const pageBox: JsonObject = {
  type: 'object',
  properties: {
    x: { type: 'number' },
    y: { type: 'number' },
    width: { type: 'number', exclusiveMinimum: 0 },
    height: { type: 'number', exclusiveMinimum: 0 },
  },
  required: ['x', 'y', 'width', 'height'],
  additionalProperties: false,
}
const CAPABILITIES: readonly AgentCapability[] = [
  ['pdf.page.rotate', 'Rotate selected pages in 90-degree increments.', false, { pages: selector, degrees: { enum: [90, 180, 270, -90, -180, -270] } }, ['pages', 'degrees']],
  ['pdf.page.insert_blank', 'Insert a blank page.', false, { at: { type: 'integer', minimum: 1 }, size: pageSize }, ['at']],
  ['pdf.page.delete', 'Permanently remove selected pages.', true, { pages: { type: 'array', minItems: 1, items: { type: 'integer', minimum: 1 }, uniqueItems: true } }, ['pages']],
  ['pdf.page.reorder', 'Rebuild the document in a requested page order.', true, { order: { type: 'array', minItems: 1, items: { type: 'integer', minimum: 1 }, uniqueItems: true } }, ['order']],
  ['pdf.page.crop', 'Replace selected page media and crop boxes.', true, { pages: selector, box: pageBox }, ['pages', 'box']],
  ['pdf.page.resize', 'Scale selected page content and page geometry.', true, { pages: selector, width: { type: 'number', exclusiveMinimum: 0 }, height: { type: 'number', exclusiveMinimum: 0 }, fit: { enum: ['stretch', 'contain'] } }, ['pages', 'width', 'height', 'fit']],
  ['pdf.page.n_up', 'Compose source pages into a new n-up document.', true, { n: { enum: [2, 4, 6, 9] }, pageSize }, ['n']],
].map(([name, description, destructive, properties, required]) => ({ name, description, destructive, requiresConfirmation: destructive, inputSchema: { type: 'object', properties, required, additionalProperties: false }, metadata: { execution: 'local' } } as AgentCapability))

async function identity(artifact: PdfAgentArtifact): Promise<AgentArtifactIdentity> { const fingerprint = await fingerprintBytes(artifact.bytes); return { artifactId: artifact.artifactId, format: 'pdf', mediaType: 'application/pdf', revision: fingerprint, fingerprint } }
function exact(input: JsonObject, fields: string[]) { for (const key of Object.keys(input)) if (!fields.includes(key)) throw new Error(`unknown input field ${key}`) }
function finite(value: unknown, name: string, positive = false): number { if (typeof value !== 'number' || !Number.isFinite(value) || positive && value <= 0) throw new Error(`${name} must be ${positive ? 'a positive ' : 'a '}finite number`); return value }
function integer(value: unknown, name: string, minimum = 1): number { if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < minimum) throw new Error(`${name} must be an integer at least ${minimum}`); return value }
function pages(value: unknown, all = true): number[] | 'all' { if (all && value === 'all') return 'all'; if (!Array.isArray(value) || value.length === 0 || value.some((item) => typeof item !== 'number' || !Number.isSafeInteger(item) || item < 1) || new Set(value).size !== value.length) throw new Error('pages must be all or a non-empty array of unique positive integers'); return [...value] as number[] }
function size(value: unknown, name: string): { width: number; height: number } { if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${name} must be an object`); const input = value as JsonObject; exact(input, ['width', 'height']); return { width: finite(input.width, `${name}.width`, true), height: finite(input.height, `${name}.height`, true) } }
function box(value: unknown): { x: number; y: number; width: number; height: number } { if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('box must be an object'); const input = value as JsonObject; exact(input, ['x', 'y', 'width', 'height']); return { x: finite(input.x, 'box.x'), y: finite(input.y, 'box.y'), width: finite(input.width, 'box.width', true), height: finite(input.height, 'box.height', true) } }
function pageOperation(operation: AgentOperation): PageOpSpec {
  const input = operation.input
  switch (operation.name) {
    case 'pdf.page.rotate': { exact(input, ['pages', 'degrees']); const degrees = input.degrees; if (![90, 180, 270, -90, -180, -270].includes(degrees as number)) throw new Error('degrees must be a supported 90-degree increment'); return { type: 'rotate', pages: pages(input.pages), degrees: degrees as 90 } }
    case 'pdf.page.insert_blank': exact(input, ['at', 'size']); return { type: 'insertBlank', at: integer(input.at, 'at'), ...(input.size === undefined ? {} : { size: size(input.size, 'size') }) }
    case 'pdf.page.delete': exact(input, ['pages']); return { type: 'delete', pages: pages(input.pages, false) as number[] }
    case 'pdf.page.reorder': exact(input, ['order']); return { type: 'reorder', order: pages(input.order, false) as number[] }
    case 'pdf.page.crop': exact(input, ['pages', 'box']); return { type: 'crop', pages: pages(input.pages), box: box(input.box) }
    case 'pdf.page.resize': { exact(input, ['pages', 'width', 'height', 'fit']); if (input.fit !== 'stretch' && input.fit !== 'contain') throw new Error('fit must be stretch or contain'); return { type: 'resize', pages: pages(input.pages), width: finite(input.width, 'width', true), height: finite(input.height, 'height', true), fit: input.fit } }
    case 'pdf.page.n_up': { exact(input, ['n', 'pageSize']); if (![2, 4, 6, 9].includes(input.n as number)) throw new Error('n must be 2, 4, 6, or 9'); return { type: 'nUp', n: input.n as 2, ...(input.pageSize === undefined ? {} : { pageSize: size(input.pageSize, 'pageSize') }) } }
    default: throw new Error(`unsupported PDF operation ${operation.name}`)
  }
}
function operations(values: readonly AgentOperation[]): { value?: PageOpSpec[]; issues: AgentIssue[] } { try { return { value: values.map(pageOperation), issues: [] } } catch (error) { return { issues: [{ ...errorIssue(error, '/operations'), severity: 'refusal' }] } } }
async function transform(artifact: PdfAgentArtifact, values: readonly AgentOperation[]): Promise<{ artifact: PdfAgentArtifact; identity: AgentArtifactIdentity; info: PdfDocumentInfo }> { const bytes = await applyPageOps(artifact.bytes, values.map(pageOperation)); const next = { kind: 'injoffice.pdf' as const, artifactId: artifact.artifactId, bytes }; return { artifact: next, identity: await identity(next), info: await readInfo(bytes) } }

export function createPdfAgentAdapter(): AgentArtifactAdapter<PdfAgentArtifact> {
  return {
    id: 'injoffice.pdf.page-operations-v1', format: 'pdf',
    supports(value): value is PdfAgentArtifact { return typeof value === 'object' && value !== null && (value as Partial<PdfAgentArtifact>).kind === 'injoffice.pdf' && typeof (value as Partial<PdfAgentArtifact>).artifactId === 'string' && (value as Partial<PdfAgentArtifact>).bytes instanceof Uint8Array },
    async identity({ artifact }) { return identity(artifact) }, async capabilities() { return structuredClone(CAPABILITIES) },
    async inspect({ artifact, cursor, maxItems, maxBytes }) { const info = await readInfo(artifact.bytes); return boundedObject({ byteLength: artifact.bytes.byteLength, pageCount: info.pageCount }, 'pages', info.pages, cursorOffset(cursor), maxItems, maxBytes) },
    async read({ artifact, query, cursor, maxItems, maxBytes }) { const info = await readInfo(artifact.bytes); const requested = query?.page; if (requested !== undefined) { if (typeof requested !== 'number' || !Number.isSafeInteger(requested) || requested < 1 || requested > info.pageCount) throw new Error('requested PDF page does not exist'); return { data: { pages: [toJson(info.pages[requested - 1])] }, itemCount: 1, truncated: false } }; return boundedObject({}, 'pages', info.pages, cursorOffset(cursor), maxItems, maxBytes) },
    async validate(request) { const issues: AgentIssue[] = []; try { assertFresh(await identity(request.artifact), request.sourceIdentity) } catch (error) { issues.push({ ...errorIssue(error), severity: 'error' }) }; const parsed = operations(request.changeSet.operations); issues.push(...parsed.issues); if (!issues.length) try { await transform(request.artifact, request.changeSet.operations) } catch (error) { issues.push({ ...errorIssue(error, '/operations'), severity: 'refusal' }) }; return issues },
    async preview(request) { assertFresh(await identity(request.artifact), request.sourceIdentity); const parsed = operations(request.changeSet.operations); if (!parsed.value) return { data: toJson({ mode: 'refused' }), issues: parsed.issues, evidence: [] }; const result = await transform(request.artifact, request.changeSet.operations); return { data: toJson({ pageCount: result.info.pageCount, pages: result.info.pages, outputFingerprint: result.identity.fingerprint, outputByteLength: result.artifact.bytes.byteLength }), issues: [], evidence: [{ kind: 'pdf.fresh-parse', data: { pageCount: result.info.pageCount } }] } },
    async diff(request) { const before = await readInfo(request.artifact.bytes); const after = await transform(request.artifact, request.changeSet.operations); return { data: { before: toJson(before), after: toJson(after.info), operations: request.changeSet.operations.map(({ operationId, name }) => ({ operationId, name })), outputFingerprint: after.identity.fingerprint }, issues: [], evidence: [] } },
    async commit(request) { assertFresh(await identity(request.artifact), request.sourceIdentity); const before = await readInfo(request.artifact.bytes); const result = await transform(request.artifact, request.changeSet.operations); return { artifact: result.artifact, identity: result.identity, data: { before: toJson(before), after: toJson(result.info), operationCount: request.changeSet.operations.length }, evidence: [{ kind: 'pdf.fresh-parse', data: { fingerprint: result.identity.fingerprint, byteLength: result.artifact.bytes.byteLength } }] } },
    async verify(request) { try { assertFresh(await identity(request.artifact), request.sourceIdentity); const info = await readInfo(request.artifact.bytes); if (request.stage === 'planned') { const parsed = operations(request.changeSet.operations); if (!parsed.value) return { verified: false, checks: [{ name: 'pdf-page-operations', passed: false }], issues: parsed.issues }; await transform(request.artifact, request.changeSet.operations) }; return { verified: true, checks: [{ name: request.stage === 'planned' ? 'pdf-transform-preview' : 'pdf-fresh-output-parse', passed: true, evidence: [{ kind: 'pdf.document-info', data: toJson(info) }] }], issues: [] } } catch (error) { return { verified: false, checks: [{ name: 'pdf-output', passed: false }], issues: [{ ...errorIssue(error), severity: 'error' }] } } },
  }
}
