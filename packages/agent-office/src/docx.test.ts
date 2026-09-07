import { readFileSync } from 'node:fs'
import { createAgentSession } from '@injoffice/agent-tools'
import type { NativeDocxDocumentV1, NativeDocxOfficeMutationEnvelopeV1 } from '@injoffice/docs/native-docx'
import { describe, expect, it } from 'vitest'
import { createDocxAgentAdapter, type DocxAgentArtifact } from './docx.js'

const record = JSON.parse(readFileSync(new URL('../../../go/officecompat/corpus/generated/expected/docx-strict-relocated.json', import.meta.url), 'utf8')) as { native: NativeDocxDocumentV1 }
function hostArtifact(): DocxAgentArtifact & { envelopes: NativeDocxOfficeMutationEnvelopeV1[] } {
  let document = structuredClone(record.native); const envelopes: NativeDocxOfficeMutationEnvelopeV1[] = []
  return { kind: 'injoffice.docx', artifactId: 'docx:1', envelopes, async snapshot() { return { document } }, async apply(envelope) { const next = structuredClone(document); const mutation = envelope.payload.mutations[0]!; const run = next.body.blocks[0]!.paragraph!.runs[0]!; if (mutation.target_id !== run.id) throw new Error('unexpected test target'); run.text = mutation.text; next.source.package_sha256 = `sha256:${'b'.repeat(64)}`; next.revision = `rev:${'b'.repeat(32)}`; document = next; envelopes.push(structuredClone(envelope)); return { snapshot: { document }, receipt: { saved: true } } } }
}

describe('DOCX agent session', () => {
  it('runs a simple semantic text replacement through the native host', async () => {
    const artifact = hostArtifact(); const run = record.native.body.blocks[0]!.paragraph!.runs[0]!
    const session = await createAgentSession({ artifact, actor: { id: 'agent-1', kind: 'agent' }, adapter: createDocxAgentAdapter(), confirmDestructive: () => true })
    expect((await session.capabilities()).capabilities[0]?.name).toBe('docx.text.replace'); expect((await session.inspect({ maxItems: 10, maxBytes: 20_000 })).identity.format).toBe('docx'); expect((await session.read({ maxItems: 2, maxBytes: 20_000 })).itemCount).toBeGreaterThan(0)
    const change = await session.plan([{ name: 'docx.text.replace', input: { targetKind: 'run', targetId: run.id, expectedText: run.text!, text: 'Strict native' } }])
    expect((await change.validate()).valid).toBe(true); expect((await change.preview()).data).toMatchObject({ mode: 'validated-envelope' }); expect((await change.diff()).data).toHaveProperty('mutations')
    const committed = await change.commit({ idempotencyKey: 'docx-save-1', confirmation: true }); expect(artifact.envelopes[0]?.mutation_id).toBe('docx-save-1'); expect(committed.verification.verified).toBe(true)
  })
  it('refuses stale expected text before host execution', async () => {
    const artifact = hostArtifact(); const run = record.native.body.blocks[0]!.paragraph!.runs[0]!; const session = await createAgentSession({ artifact, actor: { id: 'agent-1', kind: 'agent' }, adapter: createDocxAgentAdapter(), confirmDestructive: () => true })
    const change = await session.plan([{ name: 'docx.text.replace', input: { targetKind: 'run', targetId: run.id, expectedText: 'wrong', text: 'new' } }]); const report = await change.validate(); expect(report.valid).toBe(false); expect(report.issues.some(({ code }) => code === 'STALE_TARGET')).toBe(true); expect(artifact.envelopes).toHaveLength(0)
  })
})
