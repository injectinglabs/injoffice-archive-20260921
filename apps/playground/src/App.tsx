import { lazy, Suspense, useEffect, useMemo, useState, useTransition } from 'react'
import {
  applyColorScheme,
  currentColorScheme,
  persistColorScheme,
  preferredColorScheme,
  readStoredColorScheme,
  type ColorScheme,
} from './colorScheme'
import { DEMO_BY_SURFACE, DEMO_GROUPS, DEMOS, preloadDemoOnIntent, type DemoDefinition } from './demoRegistry'
import OverviewPage from './pages/OverviewPage'
import { isDocsHash, parseSurface, surfaceHref, type Surface } from './route'

const DocsApp = lazy(() => import('../../docs/src/App'))

type SidecarState = 'checking' | 'connected' | 'offline'

const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').replace(/\/$/, '')

function ToolGlyph({ demo }: { demo: DemoDefinition }) {
  return <span className={`tool-glyph tool-glyph--${demo.accent}`} aria-hidden="true">{demo.glyph}</span>
}

function ToolNavigation({ surface }: { surface: Surface }) {
  return (
    <nav className="tool-nav" aria-label="InjOffice tools">
      <a className="tool-nav-home" href={surfaceHref('overview')} aria-current={surface === 'overview' ? 'page' : undefined}>
        <span className="tool-glyph tool-glyph--ink" aria-hidden="true">⌂</span>
        <span><strong>Overview</strong><small>All capabilities</small></span>
      </a>
      {DEMO_GROUPS.map((group) => (
        <section className="tool-nav-group" key={group} aria-labelledby={`nav-${group.replaceAll(' ', '-').toLowerCase()}`}>
          <h2 id={`nav-${group.replaceAll(' ', '-').toLowerCase()}`}>{group}</h2>
          {DEMOS.filter((demo) => demo.group === group).map((demo) => {
            const warmRoute = () => { preloadDemoOnIntent(demo.surface) }
            return (
              <a
                key={demo.surface}
                href={surfaceHref(demo.surface)}
                aria-current={surface === demo.surface ? 'page' : undefined}
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

export default function App() {
  const [surface, setSurface] = useState<Surface>(() => parseSurface())
  const [sidecar, setSidecar] = useState<SidecarState>('checking')
  const [scheme, setScheme] = useState<ColorScheme>(() => currentColorScheme())
  const [isRoutePending, startRouteTransition] = useTransition()

  useEffect(() => {
    const sync = () => {
      const nextSurface = parseSurface()
      preloadDemoOnIntent(nextSurface)
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
    <div className="app-shell" data-surface={surface}>
      <a className="skip-link" href="#main-content">Skip to demo</a>
      <AppHeader sidecar={sidecar} scheme={scheme} onScheme={(next) => { persistColorScheme(next); setScheme(next) }} />
      <aside className="app-sidebar"><ToolNavigation surface={surface} /></aside>
      <main className="app-main" id="main-content" data-workbench-surface={surface}>
        <div className="route-progress" data-active={isRoutePending ? 'true' : undefined} aria-hidden="true"><i /></div>
        <span className="visually-hidden" role="status" aria-live="polite">
          {isRoutePending && requestedTitle ? `Opening ${requestedTitle}…` : ''}
        </span>
        <Suspense fallback={<div className="demo-loading app-route-loading" role="status">Opening {requestedTitle ?? 'demo'}…</div>}>
          {surface === 'overview' ? <OverviewPage sidecar={sidecar} /> : demo && DemoComponent ? (
            <>
              <header className={`page-heading page-heading--${demo.accent}`}>
                <div className="page-heading-copy">
                  <div className="page-heading-package"><ToolGlyph demo={demo} /><code>{demo.packageName}</code></div>
                  <h1>{demo.title}</h1>
                  <p>{demo.description}</p>
                </div>
                <RuntimePill demo={demo} sidecar={sidecar} />
              </header>
              <section className="demo-stage" data-accent={demo.accent} aria-label={`${demo.title} interactive demo`}>
                <DemoComponent />
              </section>
            </>
          ) : null}
        </Suspense>
      </main>
    </div>
  )
}
