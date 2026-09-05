import {
  LogLevel,
  Univer,
  type IUniverConfig,
  type Plugin,
  type PluginCtor,
} from '@univerjs/core'
import { FUniver } from '@univerjs/core/lib/facade'

type PluginConfig = ConstructorParameters<PluginCtor<Plugin>>[0]
type PluginRegistration = PluginCtor<Plugin> | [PluginCtor<Plugin>, PluginConfig]
interface UniverOssPreset { plugins: PluginRegistration[] }

/** Test-only bootstrap that deliberately accepts only explicitly supplied OSS presets. */
export function createOssUniver(options: Partial<IUniverConfig> & { presets: UniverOssPreset[] }) {
  const { presets, ...config } = options
  const univer = new Univer({ logLevel: LogLevel.WARN, ...config })
  const plugins = new Map<string, { plugin: PluginCtor<Plugin>; config?: PluginConfig }>()

  for (const preset of presets) {
    for (const registration of preset.plugins) {
      const [plugin, pluginConfig] = Array.isArray(registration)
        ? registration
        : [registration, undefined]
      plugins.delete(plugin.pluginName)
      plugins.set(plugin.pluginName, { plugin, config: pluginConfig })
    }
  }
  for (const { plugin, config: pluginConfig } of plugins.values()) {
    univer.registerPlugin(plugin, pluginConfig)
  }
  return { univer, univerAPI: FUniver.newAPI(univer) }
}
