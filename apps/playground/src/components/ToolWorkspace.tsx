import { Component, Suspense, lazy, useCallback, useEffect, useId, useMemo, useRef, useState, type ComponentType, type ErrorInfo, type ReactNode } from 'react'
import type { AgentTool } from '../route'
import { resolveToolWorkspace, TOOL_WORKSPACES, workspaceHref, workspaceNavigationHash, type ToolWorkspaceFeature } from '../toolWorkspaces'
import '../design-system/live-create-edit.css'
import './ToolWorkspace.css'

async function loadFeature(tool: AgentTool, feature: ToolWorkspaceFeature, initialHash?: string): Promise<{ default: ComponentType }> {
  if (feature.id === 'agent') { const { default: Page } = await import('../pages/AgentPage'); return { default: () => <Page fixedTool={tool} /> } }
  if (feature.id === 'collab') { const { CollabDemo: Page } = await import('../collabPage'); return { default: () => <Page fixedFormat={tool} initialHash={initialHash} /> } }
  if (feature.id === 'history') { const { default: Page } = await import('../pages/HistoryPage'); return { default: () => <Page fixedFormat={tool === 'docs' ? 'docs' : 'sheets'} /> } }
  if (tool === 'sheets' && feature.id === 'editor') { const { default: Page } = await import('../UniverEditor'); return { default: () => <div className="tool-page tool-page--flush ds"><div className="tool-page-fill"><Page /></div></div> } }
  if (tool === 'sheets' && feature.id === 'native') return import('../pages/NativeRoundTripPage')
  if (tool === 'sheets' && feature.id === 'tools') return import('../pages/SheetsToolsPage')
  switch (feature.source) {
    case 'docs': return import('../pages/DocsPage')
    case 'slides': return import('../pages/SlidesPage')
    case 'pdf': return import('../pages/PdfPage')
    case 'charts': return import('../pages/ChartsPage')
    case 'pivots': return import('../pages/PivotsPage')
    case 'shapes': return import('../pages/ShapesPage')
    case 'connectors': return import('../pages/ConnectorsPage')
    case 'formulas': return import('../pages/FormulasPage')
    case 'font-metrics': return import('../pages/FontMetricsPage')
    case 'pptx-authored': return import('../pages/PptxAuthoredPage')
    case 'pptx-native': return import('../pages/PptxNativePage')
    case 'pptx-render': return import('../pages/PptxRenderPage')
    default: throw new Error(`Unsupported workspace feature: ${feature.id}`)
  }
}

class FeatureBoundary extends Component<{ children: ReactNode; onRetry: () => void }, { error: string | null }> {
  state = { error: null as string | null }
  static getDerivedStateFromError(error: unknown) { return { error: error instanceof Error ? error.message : String(error) } }
  componentDidCatch(_error: Error, _info: ErrorInfo) { /* This failure stays local to its feature panel. */ }
  reloadPage = () => {
    if (document.querySelector('[data-demo-busy="true"]')) { window.alert('Wait for the current operation to finish before reloading.'); return }
    if (window.confirm('Reload the page? Unsaved edits and pending approvals in every open view will be lost.')) window.location.reload()
  }
  render() {
    if (this.state.error) return <div className="tool-workspace__error" data-workspace-error role="alert"><h3>This view could not open</h3><p>{this.state.error}</p><p>Other views are unchanged. Retrying resets only this failed view.</p><button type="button" data-workspace-retry className="ds-btn ds-btn--outlined" onClick={this.props.onRetry}>Retry this view</button><p>If the browser cached a failed module download, reload the page. This clears unsaved edits and pending approvals in every view.</p><button type="button" data-workspace-reload className="ds-btn ds-btn--outlined" onClick={this.reloadPage}>Reload page</button></div>
    return this.props.children
  }
}

function FeatureContent({ tool, feature, initialHash, onRetry }: { tool: AgentTool; feature: ToolWorkspaceFeature; initialHash?: string; onRetry: () => void }) {
  // Explicit retry remounts this component, creating a fresh lazy loader instead
  // of reusing React.lazy's cached rejected promise. Ordinary tab switches retain it.
  const Page = useMemo(() => lazy(() => loadFeature(tool, feature, initialHash)), [tool, feature, initialHash])
  return <FeatureBoundary onRetry={onRetry}><Suspense fallback={<div className="tool-workspace__loading" data-workspace-loading role="status">Opening {feature.label.toLowerCase()}…</div>}><Page /></Suspense></FeatureBoundary>
}

function startingFeature(tool: AgentTool, hash?: string) {
  const route = resolveToolWorkspace(hash ?? (typeof location === 'undefined' ? '' : location.hash))
  return route?.tool === tool ? route.feature : 'editor'
}

function ToolWorkspace({ tool, initialHash }: { tool: AgentTool; initialHash?: string }) {
  const definition = TOOL_WORKSPACES.find((item) => item.tool === tool)!
  const instanceId = useId()
  const entryHash = useRef(initialHash ?? (typeof location === 'undefined' ? '' : location.hash)).current
  const [active, setActive] = useState(() => startingFeature(tool, initialHash))
  const [visited, setVisited] = useState(() => [startingFeature(tool, initialHash)])
  const featureHashes = useRef<Record<string, string | undefined>>({ [startingFeature(tool, initialHash)]: resolveToolWorkspace(entryHash)?.tool === tool ? entryHash : undefined })
  const panelElements = useRef(new Map<string, HTMLElement>())
  const panelHeights = useRef<Record<string, number>>({})
  const [retryKeys, setRetryKeys] = useState<Record<string, number>>({})
  const lastInGroup = useRef<Record<string, string>>({})
  const groups = [...new Set(definition.features.map((item) => item.group))]
  const current = definition.features.find((item) => item.id === active) ?? definition.features[0]!
  const groupFeatures = definition.features.filter((item) => item.group === current.group)
  const activate = useCallback((feature: string, hash: string) => {
    // A room deep link belongs to the feature's first activation, not the
    // workspace's earlier editor visit. Never overwrite a retained session.
    if (!(feature in featureHashes.current)) {
      const route = resolveToolWorkspace(hash)
      featureHashes.current[feature] = route?.tool === tool && route.feature === feature ? hash : undefined
    }
    for (const [id, element] of panelElements.current) {
      if (!element.hidden) panelHeights.current[id] = Math.max(740, element.getBoundingClientRect().height)
    }
    setActive(feature)
    setVisited((previous) => previous.includes(feature) ? previous : [...previous, feature])
  }, [tool])
  const choose = (feature: string) => {
    const href = workspaceHref(tool, feature)
    activate(feature, href)
    if (window.location.hash !== href) {
      window.dispatchEvent(new CustomEvent('injoffice:workspace-view', { detail: { tool } }))
      window.location.hash = href
    }
  }

  useEffect(() => {
    const sync = () => {
      const route = resolveToolWorkspace(workspaceNavigationHash(window.location.hash))
      if (route?.tool === tool) activate(route.feature, window.location.hash)
    }
    window.addEventListener('hashchange', sync)
    return () => window.removeEventListener('hashchange', sync)
  }, [activate, tool])

  useEffect(() => {
    lastInGroup.current[current.group] = current.id
    if (!['editor', 'native', 'pptx-native', 'collab', 'charts', 'pptx-render'].includes(current.id)) return
    // Canvas-backed editors need their now-visible dimensions after layout.
    const frame = requestAnimationFrame(() => window.dispatchEvent(new Event('resize')))
    return () => cancelAnimationFrame(frame)
  }, [current.group, current.id])

  return <div className="tool-workspace ds" data-tool-workspace={tool} data-workspace-active={current.id}>
    <header className="tool-workspace__navigation">
      <div className="tool-workspace__tabs" role="tablist" aria-label={`${definition.title} capabilities`} onKeyDown={(event) => {
        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return
        event.preventDefault()
        const tabs = Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>('[role="tab"]'))
        const index = tabs.indexOf(document.activeElement as HTMLButtonElement)
        const next = event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : (index + (event.key === 'ArrowLeft' ? -1 : 1) + tabs.length) % tabs.length
        tabs[next]?.focus()
        tabs[next]?.click()
      }}>
        {groups.map((group, index) => {
          const selected = current.group === group
          const target = selected ? current.id : lastInGroup.current[group] ?? definition.features.find((item) => item.group === group)!.id
          return <button key={group} id={`${instanceId}-group-${index}`} type="button" role="tab" data-workspace-group={group} data-workspace-feature={target} aria-selected={selected} aria-controls={`${instanceId}-panel-${target}`} tabIndex={selected ? 0 : -1} onClick={() => choose(target)}>{group}</button>
        })}
      </div>
      <div className="tool-workspace__view">
        {groupFeatures.length > 1 ? <label htmlFor={`${instanceId}-feature`}>View<select id={`${instanceId}-feature`} data-workspace-view value={current.id} onChange={(event) => choose(event.target.value)}>{groupFeatures.map((feature) => <option key={feature.id} value={feature.id}>{feature.label}</option>)}</select></label> : <strong>{current.label}</strong>}
        <p>{current.description}</p>
      </div>
      <p className="tool-workspace__sample-note">Each view has its own sample and state; edits do not transfer between views. Opened views stay available when you switch, including pending approvals.</p>
    </header>
    {definition.features.filter((feature) => visited.includes(feature.id)).map((feature) => <section key={feature.id} ref={(element) => { if (element) panelElements.current.set(feature.id, element); else panelElements.current.delete(feature.id) }} id={`${instanceId}-panel-${feature.id}`} role="tabpanel" tabIndex={feature.id === current.id ? 0 : -1} aria-labelledby={`${instanceId}-group-${groups.indexOf(feature.group)}`} data-workspace-panel={feature.id} data-workspace-retain-layout={tool === 'sheets' && (feature.id === 'editor' || feature.id === 'collab') ? 'true' : undefined} hidden={feature.id !== current.id} inert={feature.id !== current.id} aria-hidden={feature.id !== current.id ? true : undefined} style={tool === 'sheets' && feature.id !== current.id && (feature.id === 'editor' || feature.id === 'collab') ? { height: panelHeights.current[feature.id] ?? 740 } : undefined} className="tool-workspace__panel">
      <FeatureContent key={`${feature.id}-${retryKeys[feature.id] ?? 0}`} tool={tool} feature={feature} initialHash={featureHashes.current[feature.id]} onRetry={() => setRetryKeys((previous) => ({ ...previous, [feature.id]: (previous[feature.id] ?? 0) + 1 }))} />
    </section>)}
  </div>
}

export function SheetsWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="sheets" {...props} /> }
export function DocsWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="docs" {...props} /> }
export function SlidesWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="slides" {...props} /> }
export function PdfWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="pdf" {...props} /> }
