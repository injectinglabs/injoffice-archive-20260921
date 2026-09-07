import { AgentToolsError } from './errors'
import type { AgentArtifactAdapter } from './types'

export class AgentAdapterRegistry {
  readonly #adapters = new Map<string, AgentArtifactAdapter<unknown>>()

  constructor(adapters: readonly AgentArtifactAdapter<unknown>[] = []) {
    for (const adapter of adapters) this.register(adapter)
  }

  register<TArtifact>(adapter: AgentArtifactAdapter<TArtifact>): this {
    if (!adapter.id.trim() || !adapter.format.trim()) throw new AgentToolsError('INVALID_ARGUMENT', 'Adapter id and format are required.')
    if (this.#adapters.has(adapter.id)) throw new AgentToolsError('INVALID_ARGUMENT', `Adapter ${adapter.id} is already registered.`)
    this.#adapters.set(adapter.id, adapter as AgentArtifactAdapter<unknown>)
    return this
  }

  resolve<TArtifact>(artifact: TArtifact, format?: string): AgentArtifactAdapter<TArtifact> {
    const matches = [...this.#adapters.values()].filter((adapter) => (!format || adapter.format === format) && adapter.supports(artifact))
    if (matches.length === 0) throw new AgentToolsError('INVALID_ARGUMENT', `No artifact adapter supports${format ? ` format ${format}` : ' this artifact'}.`)
    if (matches.length > 1) throw new AgentToolsError('INVALID_ARGUMENT', `Multiple artifact adapters support${format ? ` format ${format}` : ' this artifact'}; select an adapter explicitly.`)
    return matches[0] as AgentArtifactAdapter<TArtifact>
  }

  list(): readonly Readonly<{ id: string; format: string }>[] {
    return [...this.#adapters.values()].map(({ id, format }) => Object.freeze({ id, format })).sort((left, right) => left.id.localeCompare(right.id))
  }
}
