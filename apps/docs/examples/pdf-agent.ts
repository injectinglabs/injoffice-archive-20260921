import { createAgentSession } from '@injoffice/agent-tools'
import { createPdfAgentAdapter } from '@injoffice/agent-office/pdf'

export async function preparePdfRotation(bytes: Uint8Array) {
  const session = await createAgentSession({
    artifact: { kind: 'injoffice.pdf', artifactId: 'review-document', bytes: bytes.slice() },
    adapter: createPdfAgentAdapter(),
    actor: { id: 'local-proposal-agent', kind: 'agent' },
  })
  const capabilities = await session.capabilities()
  if (!capabilities.capabilities.some(capability => capability.name === 'pdf.page.rotate')) {
    throw new Error('This artifact does not support page rotation')
  }
  // Deterministic proposal; this example does not call a model.
  const change = await session.plan([
    { name: 'pdf.page.rotate', input: { pages: [1], degrees: 90 } },
  ], {
    expectedRevision: session.identity.revision,
    expectedFingerprint: session.identity.fingerprint,
  })
  const validation = await change.validate()
  if (!validation.valid) throw new Error(JSON.stringify(validation.issues))
  const preview = await change.preview()
  const diff = await change.diff()
  return {
    plan: change.envelope, preview, diff,
    // Call only from your trusted, authorized review action after displaying
    // this exact plan. Do NOT register this closure as a model-accessible tool.
    async commitFromReview(approvedChangeSetId: string, idempotencyKey: string) {
      if (approvedChangeSetId !== change.envelope.changeSetId) {
        throw new Error('Approval must match the exact reviewed change set')
      }
      const receipt = await change.commit({ idempotencyKey })
      if (!receipt.verification.verified) {
        throw new Error('Write completed, but verification failed; withhold download')
      }
      return { bytes: session.artifact.bytes.slice(), receipt }
    },
  }
}
