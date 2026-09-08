import { Suspense, useCallback, useEffect, useMemo, useRef, useState, useTransition, type MouseEvent, type RefObject } from 'react'
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
import { DEMO_BY_SURFACE, DEMO_GROUPS, DEMOS, preloadDemo, preloadDemoOnIntent, type DemoDefinition } from './demoRegistry'
import OverviewPage from './pages/OverviewPage'
import { agentHref, parseAgentTool, parseSurface, surfaceHref, AGENT_TOOLS, type Surface } from './route'
import { surfaceSectionId } from './scrollSpy'
import { createRouteLoader } from './routeLoader'

type SidecarState = 'checking' | 'connected' | 'offline'

const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').replace(/\/$/, '')

function isModifiedClick(event: MouseEvent<HTMLAnchorElement>) {
  return event.metaKey || event.ctrlKey || event.shiftKey || event.altKey || event.button !== 0
}

function ToolNavigation({ surface, hash }: { surface: Surface; hash: string }) {
  const onClick = (event: MouseEvent<HTMLAnchorElement>) => {
    if (isModifiedClick(event)) return
    const href = event.currentTarget.getAttribute('href')
    if (href && location.hash === href) event.preventDefault()
  }
  const agentTool = surface === 'agent' ? parseAgentTool(hash) : undefined
  return (
    <nav className="tool-nav" id="demo-navigation" aria-label="InjOffice tools">
      {DEMO_GROUPS.map((group) => (
        <section
          className="tool-nav-group"
          key={group}
          aria-labelledby={`nav-${group.replaceAll(' ', '-').toLowerCase()}`}
        >
          <h2 id={`nav-${group.replaceAll(' ', '-').toLowerCase()}`}>{group}</h2>
          {DEMOS.filter((demo) => demo.group === group).flatMap((demo) => {
            const warmRoute = () => { preloadDemoOnIntent(demo.surface) }
            if (demo.surface === 'agent') {
              return AGENT_TOOLS.map((item) => (
                <a
                  key={item.tool}
                  href={agentHref(item.tool)}
                  aria-label={`AI change sets, ${item.label}`}
                  aria-current={agentTool === item.tool ? 'page' : undefined}
                  onPointerEnter={warmRoute}
                  onPointerDown={warmRoute}
                  onFocus={warmRoute}
                  onClick={onClick}
                >
                  {item.label}
                </a>
              ))
            }
            return (
              <a
                key={demo.surface}
                href={surfaceHref(demo.surface)}
                aria-current={surface === demo.surface ? 'page' : undefined}
                onPointerEnter={warmRoute}
                onPointerDown={warmRoute}
                onFocus={warmRoute}
                onClick={onClick}
              >
                {demo.navTitle}
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
      <div className="app-header-inner">
        <a className="app-brand" href={surfaceHref('overview')} aria-label="InjOffice overview">
          <img className="app-logo" src={`${import.meta.env.BASE_URL}logo.svg`} alt="" width={32} height={32} />
          <span><strong>InjOffice</strong></span>
        </a>
        <div className={`sidecar-status sidecar-status--${sidecar}`} role="status">
          <i aria-hidden="true" />
          {sidecar === 'checking' ? 'Checking server' : sidecar === 'connected' ? 'Browser + server ready' : 'Browser engines ready'}
        </div>
        <div className="app-header-actions">
          <ColorSchemeToggle scheme={scheme} onScheme={onScheme} />
          <a className="github-link" href="https://github.com/injectinglabs/injoffice" target="_blank" rel="noreferrer">GitHub<span aria-hidden="true">↗</span></a>
        </div>
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
            <h2 id="source-proof-title">Guide &amp; source</h2>
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

function DemoSection({
  demo,
  sidecar,
  proofOpen,
  onOpenProof,
  proofButtonRef,
  hash,
}: {
  demo: DemoDefinition
  sidecar: SidecarState
  proofOpen: boolean
  onOpenProof: () => void
  proofButtonRef: RefObject<HTMLButtonElement | null>
  hash: string
}) {
  const [revision, setRevision] = useState(0)
  const DemoComponent = demo.component
  const agentTool = demo.surface === 'agent' ? AGENT_TOOLS.find((item) => item.tool === parseAgentTool(hash)) : undefined
  const title = agentTool?.title ?? demo.title
  const description = agentTool?.description ?? demo.description
  return (
    <article
      className="demo-section"
      id={surfaceSectionId(demo.surface)}
      data-accent={demo.accent}
      aria-labelledby={`demo-title-${demo.surface}`}
    >
      <header className={`page-heading demo-context-header demo-page-header page-heading--${demo.accent}`}>
        <div className="page-heading-copy">
          <nav className="demo-breadcrumb" aria-label="Breadcrumb">
            <a href={surfaceHref('overview')}>Showcase</a><span aria-hidden="true">/</span><span>{demo.group}</span>{agentTool ? <><span aria-hidden="true">/</span><span>{agentTool.label}</span></> : null}
          </nav>
          <div className="demo-chips" aria-label="Demo tags">
            <span className="demo-chip">{demo.group}</span>
            {agentTool ? <span className="demo-chip">{agentTool.fileType}</span> : null}
            <span className="demo-chip demo-chip--pkg">{demo.packageName}</span>
          </div>
          <div className="demo-title-line">
            <h1 id={`demo-title-${demo.surface}`} tabIndex={-1}>{title}</h1>
          </div>
          <p>{description}</p>
        </div>
        <div className="demo-context-actions">
          <button
            className="demo-reset-trigger"
            type="button"
            onClick={() => setRevision((value) => value + 1)}
          >
            Reset
          </button>
          <button
            ref={proofButtonRef}
            className="source-proof-trigger"
            type="button"
            aria-haspopup="dialog"
            aria-expanded={proofOpen}
            aria-controls="source-proof-drawer"
            onClick={onOpenProof}
          >
            Guide &amp; source
          </button>
          <a className="demo-back" href={surfaceHref('overview')}>Back</a>
        </div>
      </header>
      <section className="demo-preview" aria-label={`${title} preview`}>
        <header className="demo-preview__bar">
          <span>Preview</span>
          <RuntimePill demo={demo} sidecar={sidecar} />
        </header>
        <div className="demo-stage" data-accent={demo.accent} aria-label={`${title} interactive demo`}>
          <DemoComponent key={`${demo.surface}:${agentTool?.tool ?? 'page'}:${revision}`} />
        </div>
      </section>
    </article>
  )
}

export default function App() {
  // Start with a usable catalogue, including when a deep-link chunk is offline.
  const [route, setRoute] = useState({ surface: 'overview' as Surface, hash: surfaceHref('overview') })
  const { surface, hash } = route
  const [sidecar, setSidecar] = useState<SidecarState>('checking')
  const [scheme, setScheme] = useState<ColorScheme>(() => currentColorScheme())
  const [detailsOpen, setDetailsOpen] = useState(false)
  const [pendingRoute, setPendingRoute] = useState<typeof route | null>(null)
  const [routeError, setRouteError] = useState<typeof route | null>(null)
  const [isTransitionPending, startTransition] = useTransition()
  const isRoutePending = pendingRoute !== null || isTransitionPending
  const detailsButtonRef = useRef<HTMLButtonElement>(null)
  const previousHashRef = useRef(hash)
  const closeDetails = useCallback(() => setDetailsOpen(false), [])
  const demo = useMemo(() => surface === 'overview' ? undefined : DEMO_BY_SURFACE.get(surface), [surface])
  const requestedSurface = pendingRoute?.surface ?? surface
  const requestedTitle = requestedSurface === 'overview' ? 'Overview' : DEMO_BY_SURFACE.get(requestedSurface)?.title
  const drawerKey = surface === 'agent' ? `agent:${parseAgentTool(hash)}` : surface
  const routeLoader = useMemo(() => createRouteLoader<typeof route>({
    preload: (next) => preloadDemo(next.surface),
    pending: (next) => {
      setDetailsOpen(false)
      setRouteError(null)
      setPendingRoute(next)
    },
    ready: (next) => {
      setPendingRoute(null)
      // The shared Suspense boundary retains the prior editor while React
      // resolves the warmed lazy component; no blank intermediate stage.
      startTransition(() => setRoute(next))
    },
    failed: (next) => {
      setPendingRoute(null)
      setRouteError(next)
    },
  }), [])

  useEffect(() => {
    if (previousHashRef.current === hash) return
    previousHashRef.current = hash
    document.querySelector<HTMLElement>('.tool-nav a[aria-current="page"]')?.scrollIntoView({ block: 'nearest' })
    document.querySelector<HTMLElement>(`#${surfaceSectionId(surface)} h1, #showcase-title`)?.focus({ preventScroll: true })
  }, [surface, hash])

  useEffect(() => {
    const sync = () => {
      void routeLoader.load({ surface: parseSurface(), hash: location.hash })
    }
    window.addEventListener('hashchange', sync)
    if (!location.hash) location.hash = surfaceHref('overview')
    sync()
    return () => {
      window.removeEventListener('hashchange', sync)
      routeLoader.cancel()
    }
  }, [routeLoader])

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

  useEffect(() => {
    preloadDemoOnIntent(surface)
  }, [surface])

  return (
    <div className="app-shell ds" data-surface={surface} data-layout="univer" data-navigation="text">
      <a className="skip-link" href="#main-content">Skip to demo</a>
      <AppHeader sidecar={sidecar} scheme={scheme} onScheme={(next) => { persistColorScheme(next); setScheme(next) }} />
      <div className="app-frame">
      <aside className="app-sidebar">
        <ToolNavigation surface={surface} hash={hash} />
      </aside>
      <main className="app-main" id="main-content" data-workbench-surface={surface} aria-busy={isRoutePending}>
        <div className="route-progress" data-active={isRoutePending ? 'true' : undefined} aria-hidden="true"><i /></div>
        <span className="visually-hidden" role="status" aria-live="polite">
          {isRoutePending && requestedTitle ? `Opening ${requestedTitle}…` : ''}
        </span>
        {routeError ? <div className="route-error" role="alert">
          <p>Could not open {routeError.surface === 'overview' ? 'Overview' : DEMO_BY_SURFACE.get(routeError.surface)?.title}. Your current demo is still available. Check your connection and retry; if the demo was just updated, reload this page.</p>
          <button type="button" onClick={() => { void routeLoader.load(routeError) }}>Retry</button>
          <a href={surfaceHref('overview')}>Back to showcase</a>
        </div> : null}
        <Suspense fallback={<div className="demo-loading" role="status">Opening demo…</div>}>
        {surface === 'overview' || !demo ? (
          <OverviewPage sidecar={sidecar} />
        ) : (
          <DemoSection
            demo={demo}
            sidecar={sidecar}
            proofOpen={detailsOpen}
            onOpenProof={() => setDetailsOpen(true)}
            proofButtonRef={detailsButtonRef}
            hash={hash}
          />
        )}
        </Suspense>
        {demo ? <SourceProofDrawer key={drawerKey} demo={demo} open={detailsOpen} onClose={closeDetails} returnFocusRef={detailsButtonRef} /> : null}
      </main>
      </div>
    </div>
  )
}
