import type { PrintConfigurationSnapshotV1, PrintEvent, PrintEventName, PrintHost, PrintLayoutConfig, PrintRenderConfig, PrintSnapshot, ScreenshotRequest } from './types'
import { validatePrintConfigurationSnapshot, validatePrintLayout, validatePrintRender } from './validation'

const EMPTY_HEADER_FOOTER = Object.freeze({ topLeft: '', topCenter: '', topRight: '', bottomLeft: '', bottomCenter: '', bottomRight: '' })

export function defaultPrintLayout(subUnitId: string): PrintLayoutConfig {
  return {
    area: 'CurrentSheet',
    subUnitIds: [subUnitId],
    paperSize: 'A4',
    direction: 'Portrait',
    scale: 'FitWidth',
    customScale: 100,
    freeze: [],
    margin: 'Normal',
    maxRowsEachPage: 0,
    maxColumnsEachPage: 0,
  }
}

export function defaultPrintRender(): PrintRenderConfig {
  return { gridlines: false, headings: false, hAlign: 'Start', vAlign: 'Start', headerFooter: [], headerFooterSetting: { ...EMPTY_HEADER_FOOTER } }
}

type Listener = (event: PrintEvent) => void

/** Stateful, renderer-neutral print facade. Rendering and privileged browser APIs stay host-injected. */
export class PrintManager {
  private layout: PrintLayoutConfig
  private render: PrintRenderConfig
  private open = false
  private readonly listeners = new Set<Listener>()

  constructor(private readonly host: PrintHost, subUnitId: string, initial?: { layout?: PrintLayoutConfig; render?: PrintRenderConfig }) {
    this.layout = initial?.layout ?? defaultPrintLayout(subUnitId)
    this.render = initial?.render ?? defaultPrintRender()
    this.assertValid(this.layout, this.render)
  }

  snapshot(): PrintSnapshot {
    return structuredClone({ layout: this.layout, render: this.render, dialogOpen: this.open })
  }

  configurationSnapshot(): PrintConfigurationSnapshotV1 {
    return structuredClone({ version: 1, layout: this.layout, render: this.render })
  }

  /** Atomically replace durable settings while leaving dialog visibility
   * untouched. Invalid snapshots cannot partially update either half. */
  restoreConfiguration(snapshot: unknown): boolean {
    const result = validatePrintConfigurationSnapshot(snapshot)
    if (!result.ok) return false
    if (JSON.stringify(this.configurationSnapshot()) === JSON.stringify(result.value)) return true
    this.layout = result.value.layout
    this.render = result.value.render
    this.emit('changed', false)
    return true
  }

  onEvent(listener: Listener): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  updatePrintConfig(patch: Partial<PrintLayoutConfig>): PrintSnapshot {
    const next = { ...this.layout, ...structuredClone(patch) }
    const result = validatePrintLayout(next)
    if (!result.ok) throw new TypeError(result.issues.map((entry) => `${entry.path}: ${entry.message}`).join('; '))
    this.layout = result.value
    this.emit('changed', false)
    return this.snapshot()
  }

  updatePrintRenderConfig(patch: Partial<PrintRenderConfig>): PrintSnapshot {
    const next = { ...this.render, ...structuredClone(patch), headerFooterSetting: { ...this.render.headerFooterSetting, ...structuredClone(patch.headerFooterSetting ?? {}) } }
    const result = validatePrintRender(next)
    if (!result.ok) throw new TypeError(result.issues.map((entry) => `${entry.path}: ${entry.message}`).join('; '))
    this.render = result.value
    this.emit('changed', false)
    return this.snapshot()
  }

  openPrintDialog(): boolean {
    if (this.open) return true
    if (this.emit('before-open', true)) return false
    this.open = true
    this.emit('opened', false)
    return true
  }

  closePrintDialog(): boolean {
    if (!this.open) return true
    if (this.emit('before-cancel', true)) return false
    this.open = false
    this.emit('canceled', false)
    return true
  }

  async print(): Promise<boolean> {
    this.assertValid(this.layout, this.render)
    if (this.emit('before-confirm', true)) return false
    await this.host.print(this.snapshot())
    this.open = false
    this.emit('confirmed', false)
    return true
  }

  async getScreenshot(request: ScreenshotRequest): Promise<string | false> {
    if (!this.host.screenshot) return false
    if (!request.subUnitId) throw new TypeError('subUnitId is required')
    const probe = defaultPrintLayout(request.subUnitId)
    probe.subUnitIds = [{ id: request.subUnitId, range: request.range }]
    const result = validatePrintLayout(probe)
    if (!result.ok) throw new TypeError(result.issues.map((entry) => `${entry.path}: ${entry.message}`).join('; '))
    const value = await this.host.screenshot(structuredClone(request))
    return typeof value === 'string' && value.startsWith('data:image/') ? value : false
  }

  async saveScreenshotToClipboard(request: ScreenshotRequest): Promise<boolean> {
    if (!this.host.writeClipboardImage) return false
    const image = await this.getScreenshot(request)
    return image ? Boolean(await this.host.writeClipboardImage(image)) : false
  }

  private emit(name: PrintEventName, cancelable: boolean): boolean {
    let canceled = false
    const event: PrintEvent = {
      name,
      snapshot: this.snapshot(),
      cancelable,
      get canceled() { return canceled },
      cancel() { if (cancelable) canceled = true },
    }
    for (const listener of this.listeners) {
      try { listener(event) } catch { /* observers cannot corrupt print state transitions */ }
    }
    return canceled
  }

  private assertValid(layout: PrintLayoutConfig, render: PrintRenderConfig): void {
    const left = validatePrintLayout(layout)
    const right = validatePrintRender(render)
    if (!left.ok || !right.ok) throw new TypeError([...left.ok ? [] : left.issues, ...right.ok ? [] : right.issues].map((entry) => `${entry.path}: ${entry.message}`).join('; '))
  }
}
