import { Component, Suspense, lazy, useCallback, useEffect, useId, useMemo, useRef, useState, type ComponentType, type ErrorInfo, type ReactNode } from 'react'
import type { AgentTool } from '../route'
import { resolveToolWorkspace, TOOL_WORKSPACES, workspaceExamples, workspaceNavigationHash, type ToolWorkspaceFeature } from '../toolWorkspaces'
import '../design-system/live-create-edit.css'

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
  // of reusing React.lazy's cached rejected promise. Scrolling retains it.
  const Page = useMemo(() => lazy(() => loadFeature(tool, feature, initialHash)), [tool, feature, initialHash])
  return <FeatureBoundary onRetry={onRetry}><Suspense fallback={<div className="tool-workspace__loading" data-workspace-loading role="status">Opening {feature.label.toLowerCase()}…</div>}><Page /></Suspense></FeatureBoundary>
}

function startingFeature(tool: AgentTool, hash?: string) {
  const route = resolveToolWorkspace(hash ?? (typeof location === 'undefined' ? '' : location.hash))
  return route?.tool === tool ? route.feature : 'agent'
}

function ToolWorkspace({ tool, initialHash }: { tool: AgentTool; initialHash?: string }) {
  const definition = TOOL_WORKSPACES.find((item) => item.tool === tool)!
  const instanceId = useId()
  const entryHash = useRef(initialHash ?? (typeof location === 'undefined' ? '' : location.hash)).current
  const [active, setActive] = useState(() => startingFeature(tool, initialHash))
  const [visited, setVisited] = useState(() => [startingFeature(tool, initialHash)])
  const featureHashes = useRef<Record<string, string | undefined>>({ [startingFeature(tool, initialHash)]: resolveToolWorkspace(entryHash)?.tool === tool ? entryHash : undefined })
  const panelElements = useRef(new Map<string, HTMLElement>())
  const [retryKeys, setRetryKeys] = useState<Record<string, number>>({})
  const examples = workspaceExamples(definition)
  const current = definition.features.find((item) => item.id === active) ?? definition.features[0]!
  const activate = useCallback((feature: string, hash: string) => {
    // A room deep link belongs to the feature's first activation, not the
    // workspace's earlier editor visit. Never overwrite a retained session.
    if (!(feature in featureHashes.current)) {
      const route = resolveToolWorkspace(hash)
      featureHashes.current[feature] = route?.tool === tool && route.feature === feature ? hash : undefined
    }
    setActive(feature)
    setVisited((previous) => previous.includes(feature) ? previous : [...previous, feature])
  }, [tool])

  useEffect(() => {
    const sync = () => {
      const route = resolveToolWorkspace(workspaceNavigationHash(window.location.hash))
      if (route?.tool === tool) activate(route.feature, window.location.hash)
    }
    window.addEventListener('hashchange', sync)
    window.addEventListener('injoffice:workspace-scroll', sync)
    return () => {
      window.removeEventListener('hashchange', sync)
      window.removeEventListener('injoffice:workspace-scroll', sync)
    }
  }, [activate, tool])

  useEffect(() => {
    const observer = new IntersectionObserver(entries => {
      const visible = entries.filter(entry => entry.isIntersecting).map(entry => (entry.target as HTMLElement).dataset.workspacePanel!)
      if (!visible.length) return
      for (const feature of visible) {
        if (!(feature in featureHashes.current)) featureHashes.current[feature] = undefined
      }
      setVisited(previous => [...new Set([...previous, ...visible])])
    }, { rootMargin: '160px 0px' })
    for (const element of panelElements.current.values()) observer.observe(element)
    return () => observer.disconnect()
  }, [tool])

  useEffect(() => {
    if (!['editor', 'native', 'pptx-native', 'collab', 'charts', 'pptx-render'].includes(current.id)) return
    // Canvas-backed editors need their now-visible dimensions after layout.
    const frame = requestAnimationFrame(() => window.dispatchEvent(new Event('resize')))
    return () => cancelAnimationFrame(frame)
  }, [current.group, current.id])

  return <div className="tool-workspace ds" data-tool-workspace={tool} data-workspace-active={current.id}>
    <p className="tool-workspace__sample-note">Each view has its own sample and state; edits do not transfer between views. Examples load as you reach them and keep your changes while you explore.</p>
    {examples.map((feature) => <section key={feature.id} ref={(element) => { if (element) panelElements.current.set(feature.id, element); else panelElements.current.delete(feature.id) }} id={`example-${tool}-${feature.id}`} aria-labelledby={`${instanceId}-heading-${feature.id}`} data-workspace-panel={feature.id} className="tool-workspace__panel">
      <header className="tool-workspace__heading">
        <h3 id={`${instanceId}-heading-${feature.id}`} tabIndex={-1}>{feature.label}</h3>
        <p>{feature.description}</p>
      </header>
      <div className="tool-workspace__content">
        {visited.includes(feature.id)
          ? <FeatureContent key={`${feature.id}-${retryKeys[feature.id] ?? 0}`} tool={tool} feature={feature} initialHash={featureHashes.current[feature.id]} onRetry={() => setRetryKeys((previous) => ({ ...previous, [feature.id]: (previous[feature.id] ?? 0) + 1 }))} />
          : <div className="tool-workspace__loading" data-workspace-loading><p>This example loads as you reach it.</p><button type="button" className="ds-btn ds-btn--outlined" onClick={() => activate(feature.id, window.location.hash)}>Load {feature.label.toLowerCase()}</button></div>}
      </div>
    </section>)}
  </div>
}

export function SheetsWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="sheets" {...props} /> }
export function DocsWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="docs" {...props} /> }
export function SlidesWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="slides" {...props} /> }
export function PdfWorkspace(props: { initialHash?: string }) { return <ToolWorkspace tool="pdf" {...props} /> }
