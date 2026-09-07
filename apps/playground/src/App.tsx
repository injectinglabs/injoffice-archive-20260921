import { lazy, Suspense, useCallback, useEffect, useMemo, useRef, useState, useTransition, type RefObject } from 'react'
import {
  applyColorScheme,
  currentColorScheme,
  persistColorScheme,
  preferredColorScheme,
  readStoredColorScheme,
  type ColorScheme,
} from './colorScheme'
import GuidedRecipe from './components/GuidedRecipe'
import DemoSource from './components/DemoSource'
import { DEMO_BY_SURFACE, DEMO_GROUPS, DEMOS, preloadDemoOnIntent, type DemoDefinition } from './demoRegistry'
import OverviewPage from './pages/OverviewPage'
import { isDocsHash, parseSurface, surfaceHref, type Surface } from './route'

const DocsApp = lazy(() => import('../../docs/src/App'))

type SidecarState = 'checking' | 'connected' | 'offline'

const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').replace(/\/$/, '')

function ToolGlyph({ demo }: { demo: DemoDefinition }) {
  return <span className={`tool-glyph tool-glyph--${demo.accent}`} aria-hidden="true">{demo.glyph}</span>
}

function ToolNavigation({ surface, collapsed = false }: { surface: Surface; collapsed?: boolean }) {
  return (
    <nav className="tool-nav" id="demo-navigation" aria-label="InjOffice tools">
      <a
        className="tool-nav-home"
        href={surfaceHref('overview')}
        aria-current={surface === 'overview' ? 'page' : undefined}
        aria-label={collapsed ? 'Showcase' : undefined}
        title={collapsed ? 'Showcase' : undefined}
      >
        <span className="tool-glyph tool-glyph--ink" aria-hidden="true">⌂</span>
        <span><strong>Showcase</strong><small>Find an example</small></span>
      </a>
      {DEMO_GROUPS.map((group) => (
        <section
          className="tool-nav-group"
          key={group}
          aria-label={collapsed ? group : undefined}
          aria-labelledby={collapsed ? undefined : `nav-${group.replaceAll(' ', '-').toLowerCase()}`}
        >
          <h2 id={`nav-${group.replaceAll(' ', '-').toLowerCase()}`}>{group}</h2>
          {DEMOS.filter((demo) => demo.group === group).map((demo) => {
            const warmRoute = () => { preloadDemoOnIntent(demo.surface) }
            return (
              <a
                key={demo.surface}
                href={surfaceHref(demo.surface)}
                aria-current={surface === demo.surface ? 'page' : undefined}
                aria-label={collapsed ? demo.navTitle : undefined}
                title={collapsed ? demo.navTitle : undefined}
                onPointerEnter={warmRoute}
                onPointerDown={warmRoute}
                onFocus={warmRoute}
              >
                <ToolGlyph demo={demo} />
                <span><strong>{demo.navTitle}</strong><small>{demo.packageName}</small></span>
              </a>
            )
          })}
        </section>
      ))}
    </nav>
  )
}

function NavigationToggle({ collapsed, onToggle }: { collapsed: boolean; onToggle: () => void }) {
  const label = collapsed ? 'Expand navigation' : 'Collapse navigation'
  return (
    <div className="navigation-controls">
      <span>Demo index</span>
      <button type="button" aria-label={label} title={label} aria-expanded={!collapsed} aria-controls="demo-navigation" onClick={onToggle}>
        <span className="navigation-toggle-glyph" aria-hidden="true"><i /><i /><i /></span>
      </button>
    </div>
  )
}

function RuntimePill({ demo, sidecar }: { demo: DemoDefinition; sidecar: SidecarState }) {
  const needsSidecar = demo.runtime.includes('sidecar')
  const label = needsSidecar && sidecar === 'connected' ? 'Local sidecar connected' : demo.runtime
  return (
    <span className={`runtime-pill${needsSidecar ? ` runtime-pill--${sidecar}` : ''}`} title={`Runtime: ${label}`}>
      <i aria-hidden="true" />
      {label}
    </span>
  )
}

function ColorSchemeToggle({ scheme, onScheme }: { scheme: ColorScheme; onScheme: (next: ColorScheme) => void }) {
  return (
    <div className="scheme-toggle" role="group" aria-label="Color scheme">
      <button type="button" aria-pressed={scheme === 'light'} onClick={() => onScheme('light')}>Light</button>
      <button type="button" aria-pressed={scheme === 'dark'} onClick={() => onScheme('dark')}>Dark</button>
    </div>
  )
}

function AppHeader({ sidecar, scheme, onScheme }: { sidecar: SidecarState; scheme: ColorScheme; onScheme: (next: ColorScheme) => void }) {
  return (
    <header className="app-header">
      <a className="app-brand" href={surfaceHref('overview')} aria-label="InjOffice overview">
        <span className="brand-mark" aria-hidden="true"><i /><i /><i /></span>
        <span><strong>InjOffice</strong><small>Source-authoritative file engines</small></span>
      </a>
      <div className={`sidecar-status sidecar-status--${sidecar}`} role="status">
        <i aria-hidden="true" />
        {sidecar === 'checking' ? 'Browser engines ready · checking server' : sidecar === 'connected' ? 'Browser + server fallback ready' : 'Browser engines ready'}
      </div>
      <div className="app-header-actions">
        <ColorSchemeToggle scheme={scheme} onScheme={onScheme} />
        <a className="github-link" href="#/guides">Guides</a>
        <a className="github-link" href="https://github.com/injectinglabs/injoffice" target="_blank" rel="noreferrer">View source<span aria-hidden="true">↗</span></a>
      </div>
    </header>
  )
}

function SourceProofDrawer({
  demo,
  open,
  onClose,
  returnFocusRef,
}: {
  demo: DemoDefinition
  open: boolean
  onClose: () => void
  returnFocusRef: RefObject<HTMLButtonElement | null>
}) {
  const drawerRef = useRef<HTMLElement>(null)
  const closeRef = useRef<HTMLButtonElement>(null)

  useEffect(() => {
    if (!open) return
    const previousOverflow = document.body.style.overflow
    const background = Array.from(document.querySelectorAll<HTMLElement>('.app-header, .app-sidebar, .app-main > :not(.source-proof-layer)'))
    const inertStates = background.map((element) => element.inert)
    background.forEach((element) => { element.inert = true })
    document.body.style.overflow = 'hidden'
    closeRef.current?.focus()

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault()
        onClose()
        return
      }
      if (event.key !== 'Tab' || !drawerRef.current) return
      const controls = Array.from(drawerRef.current.querySelectorAll<HTMLElement>('a[href], button:not([disabled]), input:not([disabled]), select:not([disabled]), summary, [tabindex]:not([tabindex="-1"])')).filter((element) => element.getClientRects().length > 0)
      if (controls.length === 0) return
      const first = controls[0]
      const last = controls[controls.length - 1]
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault()
        last.focus()
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault()
        first.focus()
      }
    }

    document.addEventListener('keydown', onKeyDown)
    return () => {
      document.body.style.overflow = previousOverflow
      background.forEach((element, index) => { element.inert = inertStates[index] })
      document.removeEventListener('keydown', onKeyDown)
      returnFocusRef.current?.focus()
    }
  }, [open, onClose, returnFocusRef])

  return (
    <div className="source-proof-layer" hidden={!open}>
      <button className="source-proof-backdrop" type="button" tabIndex={-1} aria-label="Close source and proof" onClick={onClose} />
      <aside
        className="source-proof-drawer"
        id="source-proof-drawer"
        ref={drawerRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby="source-proof-title"
      >
        <header>
          <div>
            <span>{demo.packageName}</span>
            <h2 id="source-proof-title">Source &amp; proof</h2>
          </div>
          <button ref={closeRef} className="source-proof-close" type="button" aria-label="Close source and proof" onClick={onClose}>×</button>
        </header>
        <div className="source-proof-body">
          <p className="source-proof-runtime">{demo.title} · {demo.runtime}</p>
          <GuidedRecipe recipe={demo.recipe} accent={demo.accent} className="guided-recipe--drawer" />
          <DemoSource source={demo.recipe.sources[0]} />
        </div>
      </aside>
    </div>
  )
}

export default function App() {
  const [surface, setSurface] = useState<Surface>(() => parseSurface())
  const [sidecar, setSidecar] = useState<SidecarState>('checking')
  const [scheme, setScheme] = useState<ColorScheme>(() => currentColorScheme())
  const [navigationExpanded, setNavigationExpanded] = useState(false)
  const [detailsOpen, setDetailsOpen] = useState(false)
  const [demoRevision, setDemoRevision] = useState(0)
  const [isRoutePending, startRouteTransition] = useTransition()
  const detailsButtonRef = useRef<HTMLButtonElement>(null)
  const previousSurfaceRef = useRef(surface)
  const closeDetails = useCallback(() => setDetailsOpen(false), [])

  useEffect(() => {
    if (previousSurfaceRef.current === surface) return
    previousSurfaceRef.current = surface
    document.querySelector<HTMLElement>('.app-main h1')?.focus({ preventScroll: true })
  }, [surface])

  useEffect(() => {
    const sync = () => {
      const nextSurface = parseSurface()
      preloadDemoOnIntent(nextSurface)
      setDetailsOpen(false)
      startRouteTransition(() => setSurface(nextSurface))
    }
    window.addEventListener('hashchange', sync)
    if (!location.hash) location.hash = surfaceHref('overview')
    return () => window.removeEventListener('hashchange', sync)
  }, [])

  useEffect(() => {
    applyColorScheme(scheme)
  }, [scheme])

  useEffect(() => {
    const media = window.matchMedia('(prefers-color-scheme: dark)')
    const onChange = () => {
      if (readStoredColorScheme()) return
      setScheme(preferredColorScheme())
    }
    media.addEventListener('change', onChange)
    return () => media.removeEventListener('change', onChange)
  }, [])

  useEffect(() => {
    const controller = new AbortController()
    const timeout = window.setTimeout(() => controller.abort(), 1800)
    void fetch(`${API_BASE}/healthz`, { signal: controller.signal })
      .then(async (response) => {
        if (!response.ok || !response.headers.get('content-type')?.includes('application/json')) throw new Error('offline')
        const body = await response.json() as { status?: string }
        setSidecar(body.status === 'ok' ? 'connected' : 'offline')
      })
      .catch(() => setSidecar('offline'))
      .finally(() => window.clearTimeout(timeout))
    return () => { window.clearTimeout(timeout); controller.abort() }
  }, [])

  const demo = useMemo(() => surface === 'overview' ? undefined : DEMO_BY_SURFACE.get(surface), [surface])
  const DemoComponent = demo?.component
  const navigationCollapsed = surface !== 'overview' && !navigationExpanded
  const requestedSurface = parseSurface()
  const requestedTitle = requestedSurface === 'overview' ? 'Overview' : DEMO_BY_SURFACE.get(requestedSurface)?.title

  if (isDocsHash()) {
    return (
      <Suspense fallback={<div className="demo-loading" role="status">Opening guides…</div>}>
        <DocsApp />
      </Suspense>
    )
  }

  return (
    <div className="app-shell" data-surface={surface} data-navigation={navigationCollapsed ? 'collapsed' : 'expanded'}>
      <a className="skip-link" href="#main-content">Skip to demo</a>
      <AppHeader sidecar={sidecar} scheme={scheme} onScheme={(next) => { persistColorScheme(next); setScheme(next) }} />
      <aside className="app-sidebar">
        {surface !== 'overview' ? <NavigationToggle collapsed={navigationCollapsed} onToggle={() => setNavigationExpanded((expanded) => !expanded)} /> : null}
        <ToolNavigation surface={surface} collapsed={navigationCollapsed} />
      </aside>
      <main className="app-main" id="main-content" data-workbench-surface={surface}>
        <div className="route-progress" data-active={isRoutePending ? 'true' : undefined} aria-hidden="true"><i /></div>
        <span className="visually-hidden" role="status" aria-live="polite">
          {isRoutePending && requestedTitle ? `Opening ${requestedTitle}…` : ''}
        </span>
        <Suspense fallback={<div className="demo-loading app-route-loading" role="status">Opening {requestedTitle ?? 'demo'}…</div>}>
          {surface === 'overview' ? <OverviewPage sidecar={sidecar} /> : demo && DemoComponent ? (
            <>
              <header className={`page-heading demo-context-header page-heading--${demo.accent}`}>
                <div className="page-heading-copy">
                  <nav className="demo-breadcrumb" aria-label="Breadcrumb">
                    <a href={surfaceHref('overview')}>Showcase</a><span aria-hidden="true">/</span><span>{demo.group}</span>
                  </nav>
                  <div className="demo-title-line"><ToolGlyph demo={demo} /><h1 tabIndex={-1}>{demo.title}</h1><code>{demo.packageName}</code></div>
                  <p>{demo.description}</p>
                </div>
                <div className="demo-context-actions">
                  <RuntimePill demo={demo} sidecar={sidecar} />
                  <button
                    className="demo-reset-trigger"
                    type="button"
                    onClick={() => setDemoRevision((revision) => revision + 1)}
                  >
                    Reset demo
                  </button>
                  <button
                    ref={detailsButtonRef}
                    className="source-proof-trigger"
                    type="button"
                    aria-haspopup="dialog"
                    aria-expanded={detailsOpen}
                    aria-controls="source-proof-drawer"
                    onClick={() => setDetailsOpen(true)}
                  >
                    Source / proof
                  </button>
                </div>
              </header>
              <section className="demo-stage" data-accent={demo.accent} aria-label={`${demo.title} interactive demo`}>
                <DemoComponent key={`${demo.surface}:${demoRevision}`} />
              </section>
            </>
          ) : null}
        </Suspense>
        {demo ? <SourceProofDrawer key={demo.surface} demo={demo} open={detailsOpen} onClose={closeDetails} returnFocusRef={detailsButtonRef} /> : null}
      </main>
    </div>
  )
}
