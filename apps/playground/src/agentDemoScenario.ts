export type AgentDemoFormat = 'xlsx' | 'docx' | 'pptx' | 'pdf'
export type AgentDemoMode = 'safe' | 'refusal'

export type AgentDemoArtifact = {
  artifactId: string
  format: AgentDemoFormat
  name: string
  revision: string
  content: Record<string, unknown>
}

export type AgentDemoOperation = {
  operationId: string
  op: string
  target: string
  value?: unknown
  input: Record<string, unknown>
}

export type AgentDemoScenario = {
  artifact: AgentDemoArtifact
  prompt: string
  summary: string
  operations: AgentDemoOperation[]
}

const SAFE_SCENARIOS: Record<AgentDemoFormat, AgentDemoScenario> = {
  xlsx: {
    artifact: {
      artifactId: 'artifact-quarterly-plan',
      format: 'xlsx',
      name: 'quarterly-plan.xlsx',
      revision: 'rev-018',
      content: {
        sheet: { id: 'sheet-forecast', name: 'Forecast' },
        headers: ['Region', 'Plan', 'Actual', 'Confidence'],
        rows: [
          ['North', 120, 126, 'High'],
          ['South', 105, 101, 'Medium'],
          ['West', 140, 151, 'High'],
          ['East', 80, 84, 'Medium'],
        ],
      },
    },
    prompt: 'Raise the East forecast confidence after the final pipeline review.',
    summary: 'Set Forecast!D5 from Medium to High.',
    operations: [{ operationId: 'op-xlsx-001', op: 'xlsx.cell.set_value', target: 'sheet-forecast!D5', value: 'High', input: { sheetId: 'sheet-forecast', cell: { row: 4, column: 3 }, value: 'High' } }],
  },
  docx: {
    artifact: {
      artifactId: 'artifact-launch-brief',
      format: 'docx',
      name: 'northstar-launch-brief.docx',
      revision: 'rev-024',
      content: {
        title: 'Northstar launch brief',
        blocks: [
          { id: 'block-summary', kind: 'paragraph', text: 'Launch readiness is on track for the October review.' },
          { id: 'block-owner', kind: 'paragraph', text: 'Owner: Product Operations' },
          { id: 'block-risk', kind: 'paragraph', text: 'Open risk: regional enablement.' },
        ],
      },
    },
    prompt: 'Update the summary with the confirmed review date.',
    summary: 'Replace one guarded paragraph while preserving surrounding runs.',
    operations: [{ operationId: 'op-docx-001', op: 'docx.text.replace', target: 'block-summary', value: 'Launch readiness is on track for the October 18 review.', input: { targetKind: 'paragraph', targetId: 'block-summary', text: 'Launch readiness is on track for the October 18 review.', expectedText: 'Launch readiness is on track for the October review.' } }],
  },
  pptx: {
    artifact: {
      artifactId: 'artifact-board-update',
      format: 'pptx',
      name: 'board-update.pptx',
      revision: 'rev-011',
      content: {
        slide: { id: 'slide-3', number: 3 },
        title: 'Launch readiness',
        subtitle: 'Five regions prepared',
        metric: { label: 'Readiness', value: '86%' },
      },
    },
    prompt: 'Update the readiness metric using the approved score.',
    summary: 'Change one stable text shape from 86% to 91%.',
    operations: [{ operationId: 'op-pptx-001', op: 'pptx.authored.slide.update', target: 'slide-3/shape-readiness', value: '91%', input: { slideId: 'slide-3', patch: { metric: { value: '91%' } } } }],
  },
  pdf: {
    artifact: {
      artifactId: 'artifact-review-packet',
      format: 'pdf',
      name: 'review-packet.pdf',
      revision: 'rev-007',
      content: {
        pageCount: 2,
        page: { number: 2, width: 612, height: 792 },
        heading: 'Approval record',
        body: 'Operations review completed. Finance sign-off pending.',
        rotation: 0,
      },
    },
    prompt: 'Rotate page two clockwise for the review packet.',
    summary: 'Rotate page two by 90° without rewriting its page content.',
    operations: [{ operationId: 'op-pdf-001', op: 'pdf.page.rotate', target: 'page-2', value: 90, input: { pages: [2], degrees: 90 } }],
  },
}

const REFUSAL_OPERATIONS: Record<AgentDemoFormat, AgentDemoOperation> = {
  xlsx: { operationId: 'op-xlsx-refuse', op: 'macro.execute', target: 'workbook', input: {} },
  docx: { operationId: 'op-docx-refuse', op: 'tracked_changes.accept_all', target: 'document', input: {} },
  pptx: { operationId: 'op-pptx-refuse', op: 'ole_object.execute', target: 'slide-3', input: {} },
  pdf: { operationId: 'op-pdf-refuse', op: 'content.redact_unverified', target: 'page-2', input: {} },
}

export function createAgentDemoScenario(format: AgentDemoFormat, mode: AgentDemoMode): AgentDemoScenario {
  const safe = SAFE_SCENARIOS[format]
  if (mode === 'safe') return structuredClone(safe)
  return {
    artifact: structuredClone(safe.artifact),
    prompt: 'Make an unsupported or unverifiable change without confirmation.',
    summary: `Ask the ${format.toUpperCase()} adapter to perform an operation it cannot prove safe.`,
    operations: [structuredClone(REFUSAL_OPERATIONS[format])],
  }
}

export function proposeAgentDemoOperations(format: AgentDemoFormat, mode: AgentDemoMode): AgentDemoOperation[] {
  return createAgentDemoScenario(format, mode).operations
}
