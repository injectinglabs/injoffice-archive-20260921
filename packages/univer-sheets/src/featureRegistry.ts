export const INJOFFICE_FEATURE_KEYS = [
  'advancedFormula',
  'charts',
  'collaboration',
  'connectors',
  'editHistory',
  'importExport',
  'outlines',
  'pivots',
  'print',
  'rangePreprocessing',
  'shapes',
  'sparklines',
] as const

export type InjOfficeFeatureKey = (typeof INJOFFICE_FEATURE_KEYS)[number]

export interface InjOfficeFeatureOption {
  enabled?: boolean
  /** Keeps the model/service active while asking its host UI contribution to hide. */
  visible?: boolean
}

export type InjOfficeFeatureConfig = Partial<Record<InjOfficeFeatureKey, boolean | InjOfficeFeatureOption>>

export interface InjOfficeFeatureActivation {
  dispose(): void | Promise<void>
  setVisible?(visible: boolean): void | Promise<void>
}

export interface InjOfficeFeatureContext {
  key: InjOfficeFeatureKey
  visible: boolean
  serverCapabilities: ReadonlySet<string>
}

export interface InjOfficeFeatureProvider {
  /** Dynamic import or other host loader. It is never called for disabled features. */
  load(): Promise<{ activate(context: Readonly<InjOfficeFeatureContext>): InjOfficeFeatureActivation | Promise<InjOfficeFeatureActivation> }>
  dependencies?: InjOfficeFeatureKey[]
  requiredServerCapabilities?: string[]
  defaultEnabled?: boolean
  defaultVisible?: boolean
}

export type InjOfficeFeatureProviders = Partial<Record<InjOfficeFeatureKey, InjOfficeFeatureProvider>>
export type InjOfficeFeatureState = 'active' | 'disabled'

export interface InjOfficeFeatureSnapshot {
  key: InjOfficeFeatureKey
  state: InjOfficeFeatureState
  visible: boolean
}

export type InjOfficeFeatureRegistryEvent =
  | { type: 'activating' | 'activated' | 'deactivating' | 'deactivated'; key: InjOfficeFeatureKey }
  | { type: 'visibility-changed'; key: InjOfficeFeatureKey; visible: boolean }
  | { type: 'failed'; key: InjOfficeFeatureKey; error: InjOfficeFeatureRegistryError }

export type InjOfficeFeatureRegistryErrorCode =
  | 'ACTIVATION_FAILED'
  | 'DEPENDENCY_CYCLE'
  | 'DEPENDENCY_DISABLED'
  | 'INVALID_PROVIDER'
  | 'MISSING_LOADER'
  | 'SERVER_CAPABILITY_MISSING'
  | 'VISIBILITY_UNSUPPORTED'

export class InjOfficeFeatureRegistryError extends Error {
  constructor(
    public readonly code: InjOfficeFeatureRegistryErrorCode,
    message: string,
    public readonly feature: InjOfficeFeatureKey,
    options?: ErrorOptions,
  ) {
    super(message, options)
    this.name = 'InjOfficeFeatureRegistryError'
  }
}

interface ActiveFeature {
  activation: InjOfficeFeatureActivation
  visible: boolean
}

interface ResolvedFeature {
  enabled: boolean
  explicitlyDisabled: boolean
  visible: boolean
}

const keySet = new Set<string>(INJOFFICE_FEATURE_KEYS)

function normalized(config: InjOfficeFeatureConfig, providers: InjOfficeFeatureProviders): Map<InjOfficeFeatureKey, ResolvedFeature> {
  return new Map(INJOFFICE_FEATURE_KEYS.map((key) => {
    const raw = config[key]
    const provider = providers[key]
    const option = typeof raw === 'boolean' ? { enabled: raw } : raw ?? {}
    const enabled = option.enabled ?? provider?.defaultEnabled ?? false
    return [key, {
      enabled,
      explicitlyDisabled: typeof raw === 'boolean' ? !raw : raw?.enabled === false,
      visible: option.visible ?? provider?.defaultVisible ?? true,
    }]
  }))
}

function validateProviders(providers: InjOfficeFeatureProviders): void {
  for (const key of INJOFFICE_FEATURE_KEYS) {
    const provider = providers[key]
    if (!provider) continue
    if (typeof provider.load !== 'function') throw new InjOfficeFeatureRegistryError('INVALID_PROVIDER', `${key} must provide a loader`, key)
    for (const dependency of provider.dependencies ?? []) {
      if (!keySet.has(dependency) || dependency === key) throw new InjOfficeFeatureRegistryError('INVALID_PROVIDER', `${key} has invalid dependency ${dependency}`, key)
    }
    for (const capability of provider.requiredServerCapabilities ?? []) {
      if (!capability || capability.length > 128) throw new InjOfficeFeatureRegistryError('INVALID_PROVIDER', `${key} has an invalid server capability`, key)
    }
  }
}

/** Unified lazy registry for independently implemented InjOffice Sheets capabilities. */
export class InjOfficeFeatureRegistry {
  private readonly providers: InjOfficeFeatureProviders
  private readonly serverCapabilities: ReadonlySet<string>
  private readonly listeners = new Set<(event: InjOfficeFeatureRegistryEvent) => void>()
  private readonly active = new Map<InjOfficeFeatureKey, ActiveFeature>()
  private transition: Promise<void> = Promise.resolve()

  constructor(providers: InjOfficeFeatureProviders, options: { serverCapabilities?: Iterable<string> } = {}) {
    validateProviders(providers)
    this.providers = { ...providers }
    this.serverCapabilities = new Set(options.serverCapabilities)
  }

  onEvent(listener: (event: InjOfficeFeatureRegistryEvent) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  snapshot(): InjOfficeFeatureSnapshot[] {
    return INJOFFICE_FEATURE_KEYS.map((key) => {
      const active = this.active.get(key)
      return { key, state: active ? 'active' : 'disabled', visible: active?.visible ?? false }
    })
  }

  /** Apply a full configuration. Concurrent calls are serialized in call order. */
  configure(config: InjOfficeFeatureConfig): Promise<void> {
    const request = structuredClone(config)
    const next = this.transition.then(() => this.apply(request))
    this.transition = next.catch(() => undefined)
    return next
  }

  async dispose(): Promise<void> {
    await this.transition
    for (const key of [...this.active.keys()].reverse()) await this.deactivate(key)
  }

  private emit(event: InjOfficeFeatureRegistryEvent): void {
    for (const listener of this.listeners) {
      try { listener(event) } catch { /* observer isolation */ }
    }
  }

  private resolve(config: InjOfficeFeatureConfig): { desired: Map<InjOfficeFeatureKey, ResolvedFeature>; order: InjOfficeFeatureKey[] } {
    const desired = normalized(config, this.providers)
    const order: InjOfficeFeatureKey[] = []
    const visiting = new Set<InjOfficeFeatureKey>()
    const visited = new Set<InjOfficeFeatureKey>()
    const visit = (key: InjOfficeFeatureKey) => {
      if (visited.has(key)) return
      if (visiting.has(key)) throw new InjOfficeFeatureRegistryError('DEPENDENCY_CYCLE', `dependency cycle includes ${key}`, key)
      const state = desired.get(key)!
      if (!state.enabled) return
      const provider = this.providers[key]
      if (!provider) throw new InjOfficeFeatureRegistryError('MISSING_LOADER', `${key} is enabled but has no provider`, key)
      const missing = (provider.requiredServerCapabilities ?? []).find((capability) => !this.serverCapabilities.has(capability))
      if (missing) throw new InjOfficeFeatureRegistryError('SERVER_CAPABILITY_MISSING', `${key} requires server capability ${missing}`, key)
      visiting.add(key)
      for (const dependency of provider.dependencies ?? []) {
        const dependencyState = desired.get(dependency)!
        if (dependencyState.explicitlyDisabled) throw new InjOfficeFeatureRegistryError('DEPENDENCY_DISABLED', `${key} requires disabled feature ${dependency}`, key)
        dependencyState.enabled = true
        visit(dependency)
      }
      visiting.delete(key)
      visited.add(key)
      order.push(key)
    }
    for (const key of INJOFFICE_FEATURE_KEYS) visit(key)
    return { desired, order }
  }

  private async apply(config: InjOfficeFeatureConfig): Promise<void> {
    const { desired, order } = this.resolve(config)
    for (const key of order) {
      const current = this.active.get(key)
      if (current && current.visible !== desired.get(key)!.visible && !current.activation.setVisible) {
        throw new InjOfficeFeatureRegistryError('VISIBILITY_UNSUPPORTED', `${key} does not support live visibility changes`, key)
      }
    }
    const newlyActive: InjOfficeFeatureKey[] = []
    try {
      for (const key of order) {
        if (this.active.has(key)) continue
        this.emit({ type: 'activating', key })
        try {
          const loaded = await this.providers[key]!.load()
          if (!loaded || typeof loaded.activate !== 'function') throw new TypeError('loader did not return an activatable feature')
          const state = desired.get(key)!
          const activation = await loaded.activate(Object.freeze({ key, visible: state.visible, serverCapabilities: new Set(this.serverCapabilities) }))
          if (!activation || typeof activation.dispose !== 'function') throw new TypeError('feature did not return a disposable activation')
          this.active.set(key, { activation, visible: state.visible })
          newlyActive.push(key)
          this.emit({ type: 'activated', key })
        } catch (cause) {
          const error = new InjOfficeFeatureRegistryError('ACTIVATION_FAILED', `${key} activation failed`, key, { cause })
          this.emit({ type: 'failed', key, error })
          throw error
        }
      }
    } catch (error) {
      for (const key of newlyActive.reverse()) await this.deactivate(key)
      throw error
    }

    const visibilityChanged: InjOfficeFeatureKey[] = []
    let visibilityFailure: InjOfficeFeatureKey | undefined
    try {
      for (const key of order) {
        const current = this.active.get(key)!
        const visible = desired.get(key)!.visible
        if (current.visible === visible) continue
        const previous = current.visible
        try {
          await current.activation.setVisible!(visible)
        } catch (cause) {
          visibilityFailure = key
          try { await current.activation.setVisible!(previous) } catch { /* best-effort host rollback */ }
          throw cause
        }
        current.visible = visible
        visibilityChanged.push(key)
        this.emit({ type: 'visibility-changed', key, visible })
      }
    } catch (cause) {
      for (const key of visibilityChanged.reverse()) {
        const current = this.active.get(key)
        if (!current) continue
        const previous = !current.visible
        try {
          await current.activation.setVisible?.(previous)
          current.visible = previous
          this.emit({ type: 'visibility-changed', key, visible: previous })
        } catch { /* retain fail-closed transition error */ }
      }
      for (const key of newlyActive.reverse()) await this.deactivate(key)
      const feature = visibilityFailure ?? order[0]
      const error = new InjOfficeFeatureRegistryError('ACTIVATION_FAILED', `${feature} visibility change failed`, feature, { cause })
      this.emit({ type: 'failed', key: feature, error })
      throw error
    }

    const keep = new Set(order)
    for (const key of [...this.active.keys()].reverse()) if (!keep.has(key)) await this.deactivate(key)
  }

  private async deactivate(key: InjOfficeFeatureKey): Promise<void> {
    const current = this.active.get(key)
    if (!current) return
    this.emit({ type: 'deactivating', key })
    await current.activation.dispose()
    this.active.delete(key)
    this.emit({ type: 'deactivated', key })
  }
}
