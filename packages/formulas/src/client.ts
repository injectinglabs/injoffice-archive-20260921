import { TARGET_FUNCTIONS } from './target.js'

/** Deliberately structural so the adapter can consume the pinned Univer facade without bundling Univer. */
export type ClientFormulaFunction = (...args: any[]) => any

export interface DisposableLike {
  dispose(): void
}

export interface UniverFormulaFacadeLike {
  registerFunction(name: string, calculate: ClientFormulaFunction, description?: string): DisposableLike
  registerAsyncFunction?(name: string, calculate: ClientFormulaFunction, description?: string): DisposableLike
  executeCalculation(): void
  stopCalculation(): void
  onCalculationResultApplied?(timeout?: number): Promise<void>
}

export interface UniverApiWithFormula {
  getFormula(): UniverFormulaFacadeLike
}

export interface FormulaRangeLike {
  setValue(value: { f: string }): unknown
}

export interface ClientFormulaDefinition {
  readonly name: string
  readonly calculate: ClientFormulaFunction
  readonly description?: string
  readonly async?: boolean
  readonly requiredCapabilities?: readonly string[]
}

export type MissingFormulaCapabilityBehavior = 'error' | 'omit'

export interface ClientFormulaConfig {
  /** Disabled activation performs no facade registration and refuses worksheet operations. */
  readonly enabled?: boolean
  /** A deterministic allow-list for host functions. Unknown names fail closed. */
  readonly enabledFunctions?: readonly string[]
  readonly capabilities?: Iterable<string>
  readonly missingCapability?: MissingFormulaCapabilityBehavior
  /** Built-in audited target names cannot be replaced unless explicitly authorized. */
  readonly allowTargetOverride?: boolean
}

export type ClientFormulaErrorCode =
  | 'INVALID_CONFIG'
  | 'DUPLICATE_FUNCTION'
  | 'TARGET_OVERRIDE_REFUSED'
  | 'CAPABILITY_MISSING'
  | 'FACADE_UNSUPPORTED'
  | 'REGISTRATION_FAILED'
  | 'DISPOSAL_FAILED'
  | 'INACTIVE'

export class ClientFormulaError extends Error {
  override readonly name = 'ClientFormulaError'

  constructor(
    readonly code: ClientFormulaErrorCode,
    message: string,
    readonly functionName?: string,
    options?: ErrorOptions,
  ) {
    super(message, options)
  }
}

export type ClientFormulaEvent =
  | { readonly type: 'activated'; readonly registered: readonly string[]; readonly omitted: readonly string[] }
  | { readonly type: 'calculation-started' }
  | { readonly type: 'calculation-completed' }
  | { readonly type: 'calculation-stopped' }
  | { readonly type: 'disposed'; readonly registered: readonly string[] }

export interface ClientFormulaActivation {
  readonly enabled: boolean
  readonly registeredFunctions: readonly string[]
  readonly omittedFunctions: readonly string[]
  readonly targetVocabulary: readonly string[]
  setFormula(range: FormulaRangeLike, formula: string): void
  recalculate(timeoutMs?: number): Promise<void>
  stopCalculation(): void
  dispose(): void
}

export interface AdvancedFormulaFeatureContext {
  readonly serverCapabilities: ReadonlySet<string>
}

export interface AdvancedFormulaFeatureProvider {
  readonly defaultEnabled: boolean
  load(): Promise<{
    activate(context: AdvancedFormulaFeatureContext): ClientFormulaActivation
  }>
}

const TARGET_VOCABULARY = Object.freeze(TARGET_FUNCTIONS.map(({ name }) => name))
const TARGET_NAMES = new Set(TARGET_VOCABULARY)

/**
 * Deterministic host-function registration plus a small worksheet/calculation
 * facade. Built-in calculation remains Univer OSS's responsibility.
 */
export class ClientFormulaIntegration {
  private readonly listeners = new Set<(event: ClientFormulaEvent) => void>()
  private readonly definitions: readonly NormalizedDefinition[]

  constructor(
    private readonly facade: UniverFormulaFacadeLike,
    definitions: readonly ClientFormulaDefinition[] = [],
  ) {
    if (!facade || typeof facade.registerFunction !== 'function' || typeof facade.executeCalculation !== 'function' || typeof facade.stopCalculation !== 'function') {
      throw new ClientFormulaError('FACADE_UNSUPPORTED', 'A Univer formula facade with registration and calculation controls is required.')
    }
    this.definitions = normalizeDefinitions(definitions)
  }

  static fromUniver(univerAPI: UniverApiWithFormula, definitions: readonly ClientFormulaDefinition[] = []): ClientFormulaIntegration {
    if (!univerAPI || typeof univerAPI.getFormula !== 'function') {
      throw new ClientFormulaError('FACADE_UNSUPPORTED', 'A Univer API exposing getFormula() is required.')
    }
    return new ClientFormulaIntegration(univerAPI.getFormula(), definitions)
  }

  onEvent(listener: (event: ClientFormulaEvent) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  get targetVocabulary(): readonly string[] {
    return TARGET_VOCABULARY
  }

  supportsAuditedTarget(name: string): boolean {
    return TARGET_NAMES.has(normalizeName(name))
  }

  activate(config: ClientFormulaConfig = {}): ClientFormulaActivation {
    const normalized = normalizeConfig(config, this.definitions)
    if (!normalized.enabled) return this.activation([], [], false)

    const selected = this.definitions.filter((definition) => normalized.enabledNames === undefined || normalized.enabledNames.has(definition.name))
    const eligible: NormalizedDefinition[] = []
    const omitted: string[] = []
    for (const definition of selected) {
      if (TARGET_NAMES.has(definition.name) && !normalized.allowTargetOverride) {
        throw new ClientFormulaError('TARGET_OVERRIDE_REFUSED', `${definition.name} is an audited built-in target; overriding it requires allowTargetOverride.`, definition.name)
      }
      const missing = definition.requiredCapabilities.find((capability) => !normalized.capabilities.has(capability))
      if (missing) {
        if (normalized.missingCapability === 'omit') {
          omitted.push(definition.name)
          continue
        }
        throw new ClientFormulaError('CAPABILITY_MISSING', `${definition.name} requires capability ${missing}.`, definition.name)
      }
      if (definition.async && typeof this.facade.registerAsyncFunction !== 'function') {
        throw new ClientFormulaError('FACADE_UNSUPPORTED', `${definition.name} requires registerAsyncFunction().`, definition.name)
      }
      eligible.push(definition)
    }

    const registrations: Array<{ name: string; disposable: DisposableLike }> = []
    try {
      for (const definition of eligible) {
        const disposable = definition.async
          ? this.facade.registerAsyncFunction!(definition.name, definition.calculate, definition.description)
          : this.facade.registerFunction(definition.name, definition.calculate, definition.description)
        if (!disposable || typeof disposable.dispose !== 'function') {
          throw new TypeError('Formula registration did not return a disposable.')
        }
        registrations.push({ name: definition.name, disposable })
      }
    } catch (cause) {
      disposeReverse(registrations)
      const name = eligible[registrations.length]?.name
      throw new ClientFormulaError('REGISTRATION_FAILED', `Failed to register${name ? ` ${name}` : ' formula function'}.`, name, { cause })
    }
    return this.activation(registrations, omitted, true)
  }

  private activation(
    registrations: Array<{ name: string; disposable: DisposableLike }>,
    omitted: readonly string[],
    enabled: boolean,
  ): ClientFormulaActivation {
    const registered = Object.freeze(registrations.map(({ name }) => name))
    const omittedNames = Object.freeze([...omitted])
    let active = enabled
    if (enabled) this.emit({ type: 'activated', registered, omitted: omittedNames })
    const assertActive = () => {
      if (!active) throw new ClientFormulaError('INACTIVE', 'The advanced formula client activation is disabled or disposed.')
    }
    return Object.freeze({
      enabled,
      registeredFunctions: registered,
      omittedFunctions: omittedNames,
      targetVocabulary: TARGET_VOCABULARY,
      setFormula: (range: FormulaRangeLike, formula: string) => {
        assertActive()
        if (!range || typeof range.setValue !== 'function') throw new ClientFormulaError('INVALID_CONFIG', 'A worksheet range exposing setValue() is required.')
        if (typeof formula !== 'string' || !formula.startsWith('=') || formula.length > 32_768) {
          throw new ClientFormulaError('INVALID_CONFIG', 'Formula must start with = and contain at most 32768 characters.')
        }
        range.setValue({ f: formula })
      },
      recalculate: async (timeoutMs = 30_000) => {
        assertActive()
        if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0 || timeoutMs > 600_000) {
          throw new ClientFormulaError('INVALID_CONFIG', 'Calculation timeout must be an integer from 1 to 600000 ms.')
        }
        if (typeof this.facade.onCalculationResultApplied !== 'function') {
          throw new ClientFormulaError('FACADE_UNSUPPORTED', 'The formula facade does not expose onCalculationResultApplied().')
        }
        this.emit({ type: 'calculation-started' })
        const completed = this.facade.onCalculationResultApplied(timeoutMs)
        this.facade.executeCalculation()
        await completed
        assertActive()
        this.emit({ type: 'calculation-completed' })
      },
      stopCalculation: () => {
        assertActive()
        this.facade.stopCalculation()
        this.emit({ type: 'calculation-stopped' })
      },
      dispose: () => {
        if (!active) return
        active = false
        const failures = disposeReverse(registrations)
        this.emit({ type: 'disposed', registered })
        if (failures.length > 0) {
          throw new ClientFormulaError('DISPOSAL_FAILED', `Failed to unregister ${failures.join(', ')}.`)
        }
      },
    })
  }

  private emit(event: ClientFormulaEvent): void {
    for (const listener of this.listeners) {
      try {
        listener(event)
      } catch {
        // Observers cannot alter calculation or registration outcomes.
      }
    }
  }
}

/** Provider compatible with `InjOfficeFeatureRegistry`'s advancedFormula slot. */
export function createAdvancedFormulaFeatureProvider(
  integration: ClientFormulaIntegration,
  config: Omit<ClientFormulaConfig, 'capabilities'> = {},
): AdvancedFormulaFeatureProvider {
  return Object.freeze({
    defaultEnabled: true,
    async load() {
      return {
        activate(context: AdvancedFormulaFeatureContext) {
          return integration.activate({ ...config, capabilities: context.serverCapabilities })
        },
      }
    },
  })
}

interface NormalizedDefinition {
  readonly name: string
  readonly calculate: ClientFormulaFunction
  readonly description?: string
  readonly async: boolean
  readonly requiredCapabilities: readonly string[]
}

interface NormalizedConfig {
  readonly enabled: boolean
  readonly enabledNames?: ReadonlySet<string>
  readonly capabilities: ReadonlySet<string>
  readonly missingCapability: MissingFormulaCapabilityBehavior
  readonly allowTargetOverride: boolean
}

function normalizeDefinitions(input: readonly ClientFormulaDefinition[]): readonly NormalizedDefinition[] {
  if (!Array.isArray(input)) throw new ClientFormulaError('INVALID_CONFIG', 'Formula definitions must be an array.')
  const seen = new Set<string>()
  const result = input.map((definition, index) => {
    if (!definition || typeof definition !== 'object' || typeof definition.calculate !== 'function') {
      throw new ClientFormulaError('INVALID_CONFIG', `Formula definition ${index} must provide calculate().`)
    }
    const name = normalizeName(definition.name)
    if (!/^[A-Z_][A-Z0-9_.]*$/.test(name) || name.length > 255) {
      throw new ClientFormulaError('INVALID_CONFIG', `Formula function name ${JSON.stringify(definition.name)} is invalid.`, name)
    }
    if (seen.has(name)) throw new ClientFormulaError('DUPLICATE_FUNCTION', `Formula function ${name} is duplicated.`, name)
    seen.add(name)
    if (definition.description !== undefined && (typeof definition.description !== 'string' || definition.description.length > 4096)) {
      throw new ClientFormulaError('INVALID_CONFIG', `${name} description must contain at most 4096 characters.`, name)
    }
    if (definition.requiredCapabilities !== undefined && !Array.isArray(definition.requiredCapabilities)) {
      throw new ClientFormulaError('INVALID_CONFIG', `${name} requiredCapabilities must be an array.`, name)
    }
    const requiredCapabilities = normalizeCapabilities(definition.requiredCapabilities ?? [], `${name} requiredCapabilities`)
    return Object.freeze({
      name,
      calculate: definition.calculate,
      description: definition.description,
      async: definition.async === true,
      requiredCapabilities,
    })
  })
  result.sort((left, right) => compareText(left.name, right.name))
  return Object.freeze(result)
}

function normalizeConfig(config: ClientFormulaConfig, definitions: readonly NormalizedDefinition[]): NormalizedConfig {
  if (!config || typeof config !== 'object') throw new ClientFormulaError('INVALID_CONFIG', 'Formula client config must be an object.')
  const missingCapability = config.missingCapability ?? 'error'
  if (missingCapability !== 'error' && missingCapability !== 'omit') {
    throw new ClientFormulaError('INVALID_CONFIG', 'missingCapability must be error or omit.')
  }
  let enabledNames: Set<string> | undefined
  if (config.enabledFunctions !== undefined) {
    if (!Array.isArray(config.enabledFunctions)) throw new ClientFormulaError('INVALID_CONFIG', 'enabledFunctions must be an array.')
    enabledNames = new Set<string>()
    const known = new Set(definitions.map(({ name }) => name))
    for (const raw of config.enabledFunctions) {
      const name = normalizeName(raw)
      if (!known.has(name)) throw new ClientFormulaError('INVALID_CONFIG', `Unknown configured formula function ${name}.`, name)
      if (enabledNames.has(name)) throw new ClientFormulaError('INVALID_CONFIG', `Configured formula function ${name} is duplicated.`, name)
      enabledNames.add(name)
    }
  }
  return {
    enabled: config.enabled !== false,
    enabledNames,
    capabilities: new Set(normalizeCapabilities(config.capabilities ?? [], 'capabilities')),
    missingCapability,
    allowTargetOverride: config.allowTargetOverride === true,
  }
}

function normalizeCapabilities(values: Iterable<string>, label: string): readonly string[] {
  const seen = new Set<string>()
  if (typeof values === 'string' || values == null) {
    throw new ClientFormulaError('INVALID_CONFIG', `${label} must be an iterable collection, not text.`)
  }
  try {
    for (const value of values) {
      if (typeof value !== 'string' || !/^[a-z][a-z0-9._:-]*$/.test(value) || value.length > 128) {
        throw new ClientFormulaError('INVALID_CONFIG', `${label} contains an invalid capability.`)
      }
      seen.add(value)
    }
  } catch (cause) {
    if (cause instanceof ClientFormulaError) throw cause
    throw new ClientFormulaError('INVALID_CONFIG', `${label} must be iterable.`, undefined, { cause })
  }
  return [...seen].sort(compareText)
}

function disposeReverse(registrations: Array<{ name: string; disposable: DisposableLike }>): string[] {
  const failures: string[] = []
  for (const registration of [...registrations].reverse()) {
    try {
      registration.disposable.dispose()
    } catch {
      failures.push(registration.name)
    }
  }
  return failures
}

function normalizeName(value: unknown): string {
  return typeof value === 'string' ? value.trim().toUpperCase() : ''
}

function compareText(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}
