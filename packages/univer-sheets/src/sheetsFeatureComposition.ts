import {
  INJOFFICE_FEATURE_KEYS,
  InjOfficeFeatureRegistry,
  type InjOfficeFeatureConfig,
  type InjOfficeFeatureKey,
  type InjOfficeFeatureProviders,
  type InjOfficeFeatureSnapshot,
} from './featureRegistry'
import {
  createSheetsPresetBundle,
  DEFAULT_SHEETS_FEATURES,
  SHEETS_FEATURE_KEYS,
  type SheetsFeatureConfig,
  type SheetsFeatureKey,
  type SheetsPresetBundleOptions,
} from './sheetsFeaturePresets'

export const INJOFFICE_SHEETS_FEATURE_KEYS = [
  ...SHEETS_FEATURE_KEYS,
  ...INJOFFICE_FEATURE_KEYS,
] as const

export type InjOfficeSheetsFeatureKey = (typeof INJOFFICE_SHEETS_FEATURE_KEYS)[number]

/** One configuration object for static Univer OSS presets and runtime InjOffice features. */
export type InjOfficeSheetsFeatureConfig = SheetsFeatureConfig & InjOfficeFeatureConfig

export interface InjOfficeSheetsFeatureCompositionOptions extends Omit<SheetsPresetBundleOptions, 'features'> {
  features?: InjOfficeSheetsFeatureConfig
  providers?: InjOfficeFeatureProviders
  serverCapabilities?: Iterable<string>
}

export interface InjOfficeSheetsFeatureSnapshot {
  key: InjOfficeSheetsFeatureKey
  source: 'univer-oss-preset' | 'injoffice-provider'
  state: 'included' | 'active' | 'disabled'
  visible: boolean
}

function own(value: object, key: PropertyKey): boolean {
  return Object.prototype.hasOwnProperty.call(value, key)
}

function validateConfig(config: InjOfficeSheetsFeatureConfig): void {
  if (!config || typeof config !== 'object' || Array.isArray(config)) throw new TypeError('Sheets feature configuration must be an object')
  const known = new Set<string>(INJOFFICE_SHEETS_FEATURE_KEYS)
  for (const [key, value] of Object.entries(config)) {
    if (!known.has(key)) throw new TypeError(`Unknown Sheets feature ${key}`)
    if (SHEETS_FEATURE_KEYS.includes(key as SheetsFeatureKey)) {
      if (typeof value !== 'boolean') throw new TypeError(`${key} must be a boolean because Univer preset UI and model registration are inseparable`)
      continue
    }
    if (typeof value === 'boolean') continue
    if (!value || typeof value !== 'object' || Array.isArray(value)
      || Object.keys(value).some((option) => option !== 'enabled' && option !== 'visible')
      || own(value, 'enabled') && typeof value.enabled !== 'boolean'
      || own(value, 'visible') && typeof value.visible !== 'boolean') {
      throw new TypeError(`${key} must be a boolean or an enabled/visible option`)
    }
  }
}

function splitConfig(config: InjOfficeSheetsFeatureConfig): {
  presets: SheetsFeatureConfig
  advanced: InjOfficeFeatureConfig
} {
  validateConfig(config)
  const presets: SheetsFeatureConfig = {}
  const advanced: InjOfficeFeatureConfig = {}
  for (const key of SHEETS_FEATURE_KEYS) if (own(config, key)) presets[key] = config[key]
  for (const key of INJOFFICE_FEATURE_KEYS) if (own(config, key)) advanced[key] = config[key]
  return { presets, advanced }
}

/**
 * Deterministically composes Univer OSS presets and lazy InjOffice providers.
 * Presets are fixed once this object is constructed because Univer has no safe
 * public plugin-unregistration API; advanced providers remain configurable.
 */
export class InjOfficeSheetsFeatureComposition {
  readonly bundle: ReturnType<typeof createSheetsPresetBundle>
  readonly registry: InjOfficeFeatureRegistry
  private readonly presetFeatures: Readonly<Record<SheetsFeatureKey, boolean>>
  private readonly initialAdvanced: InjOfficeFeatureConfig

  constructor(options: InjOfficeSheetsFeatureCompositionOptions) {
    if (!options || typeof options !== 'object') throw new TypeError('Sheets feature composition options are required')
    const { presets, advanced } = splitConfig(options.features ?? {})
    this.bundle = createSheetsPresetBundle({
      container: options.container,
      ...(options.workerURL === undefined ? {} : { workerURL: options.workerURL }),
      features: presets,
    })
    this.presetFeatures = this.bundle.features
    this.initialAdvanced = structuredClone(advanced)
    this.registry = new InjOfficeFeatureRegistry(options.providers ?? {}, {
      serverCapabilities: options.serverCapabilities,
    })
  }

  activate(): Promise<void> {
    return this.registry.configure(this.initialAdvanced)
  }

  /** Apply a full advanced-feature configuration while guarding static presets. */
  configure(features: InjOfficeSheetsFeatureConfig): Promise<void> {
    const { presets, advanced } = splitConfig(features)
    for (const key of SHEETS_FEATURE_KEYS) {
      if (own(presets, key) && presets[key] !== this.presetFeatures[key]) {
        return Promise.reject(new TypeError(`${key} is a static Univer preset and cannot change after composition`))
      }
    }
    return this.registry.configure(advanced)
  }

  snapshot(): InjOfficeSheetsFeatureSnapshot[] {
    const advanced = new Map<InjOfficeFeatureKey, InjOfficeFeatureSnapshot>(
      this.registry.snapshot().map((entry) => [entry.key, entry]),
    )
    return INJOFFICE_SHEETS_FEATURE_KEYS.map((key) => {
      if (SHEETS_FEATURE_KEYS.includes(key as SheetsFeatureKey)) {
        const included = this.presetFeatures[key as SheetsFeatureKey]
        return { key, source: 'univer-oss-preset', state: included ? 'included' : 'disabled', visible: included }
      }
      const state = advanced.get(key as InjOfficeFeatureKey)!
      return { key, source: 'injoffice-provider', state: state.state, visible: state.visible }
    })
  }

  dispose(): Promise<void> {
    return this.registry.dispose()
  }
}

export function createInjOfficeSheetsFeatureComposition(
  options: InjOfficeSheetsFeatureCompositionOptions,
): InjOfficeSheetsFeatureComposition {
  return new InjOfficeSheetsFeatureComposition(options)
}

/** Defaults exposed for configuration UIs without instantiating Univer. */
export const DEFAULT_INJOFFICE_SHEETS_PRESET_FEATURES = DEFAULT_SHEETS_FEATURES
