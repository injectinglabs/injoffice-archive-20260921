import { AgentAdapterRegistry, type AgentArtifactAdapter } from '@injoffice/agent-tools'

export { AgentAdapterRegistry }
export type { AgentArtifactAdapter, AgentArtifactIdentity, AgentCapability, AgentIssue } from '@injoffice/agent-tools'

/** Format-neutral registry helper; callers opt into only the subpath adapters they use. */
export function createOfficeAdapterRegistry(adapters: readonly AgentArtifactAdapter<unknown>[] = []): AgentAdapterRegistry {
  return new AgentAdapterRegistry(adapters)
}
