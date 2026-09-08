export type MockAgentProposal = { operations: Array<{
  name: string; operationId: string; input: Record<string, unknown>
}> }

export class MockAgentProposalError extends Error {
  readonly status: 400 | 422
  constructor(status: 400 | 422, message: string) { super(message); this.status = status; this.name = 'MockAgentProposalError' }
}
