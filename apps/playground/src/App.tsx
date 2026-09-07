import { lazy, Suspense, useCallback, useEffect, useMemo, useRef, useState, type MouseEvent, type RefObject } from 'react'
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
import { DsMark } from './design-system/primitives'
import { DEMO_BY_SURFACE, DEMO_GROUPS, DEMOS, preloadDemo, preloadDemoOnIntent, type DemoDefinition } from './demoRegistry'
import OverviewPage from './pages/OverviewPage'
import { agentHref, isDesignSystemHash, isDocsHash, parseAgentTool, parseSurface, surfaceHref, AGENT_TOOLS, type Surface } from './route'
import { surfaceSectionId } from './scrollSpy'

const DocsApp = lazy(() => import('../../docs/src/App'))
const DesignSystemGallery = lazy(() => import('./design-system/GalleryPage'))

type SidecarState = 'checking' | 'connected' | 'offline'

const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').replace(/\/$/, '')

function isModifiedClick(event: MouseEvent<HTMLAnchorElement>) {
  return event.metaKey || event.ctrlKey || event.shiftKey || event.altKey || event.button !== 0
}

function ToolNavigation({ surface }: { surface: Surface }) {
  const onClick = (event: MouseEvent<HTMLAnchorElement>) => {
    if (isModifiedClick(event)) return
    const href = event.currentTarget.getAttribute('href')
    if (href && location.hash === href) event.preventDefault()
  }
  const agentTool = surface === 'agent' ? parseAgentTool() : undefined
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
          <DsMark />
          <span><strong>InjOffice</strong></span>
        </a>
        <nav className="app-header-links" aria-label="Site">
          <a href={surfaceHref('overview')} aria-current="page">Showcase</a>
          <a href="#/guides">Guides</a>
          <a href="#/design-system">Design system</a>
        </nav>
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

function DemoSection({
  demo,
  sidecar,
  proofOpen,
  onOpenProof,
  proofButtonRef,
}: {
  demo: DemoDefinition
  sidecar: SidecarState
  proofOpen: boolean
  onOpenProof: () => void
  proofButtonRef: RefObject<HTMLButtonElement | null>
}) {
  const [revision, setRevision] = useState(0)
  const DemoComponent = demo.component
  const agentTool = demo.surface === 'agent' ? AGENT_TOOLS.find((item) => item.tool === parseAgentTool()) : undefined
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
            Source
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
          <Suspense fallback={<div className="demo-loading" role="status">Opening {title}…</div>}>
            <DemoComponent key={`${demo.surface}:${agentTool?.tool ?? 'page'}:${revision}`} />
          </Suspense>
        </div>
      </section>
    </article>
  )
}

export default function App() {
  const [surface, setSurface] = useState<Surface>(() => parseSurface())
  const [sidecar, setSidecar] = useState<SidecarState>('checking')
  const [scheme, setScheme] = useState<ColorScheme>(() => currentColorScheme())
  const [detailsOpen, setDetailsOpen] = useState(false)
  const [isRoutePending, setIsRoutePending] = useState(false)
  const [, setHashTick] = useState(0)
  const detailsButtonRef = useRef<HTMLButtonElement>(null)
  const surfaceRef = useRef(surface)
  const previousSurfaceRef = useRef(surface)
  const closeDetails = useCallback(() => setDetailsOpen(false), [])
  const demo = useMemo(() => surface === 'overview' ? undefined : DEMO_BY_SURFACE.get(surface), [surface])
  const requestedTitle = surface === 'overview' ? 'Overview' : demo?.title

  useEffect(() => {
    surfaceRef.current = surface
  }, [surface])

  useEffect(() => {
    if (previousSurfaceRef.current === surface) return
    previousSurfaceRef.current = surface
    document.querySelector<HTMLElement>('.tool-nav a[aria-current="page"]')?.scrollIntoView({ block: 'nearest' })
    document.querySelector<HTMLElement>(`#${surfaceSectionId(surface)} h1, #showcase-title`)?.focus({ preventScroll: true })
  }, [surface])

  useEffect(() => {
    const sync = () => {
      setHashTick((tick) => tick + 1)
      if (isDocsHash() || isDesignSystemHash()) return
      const nextSurface = parseSurface()
      preloadDemoOnIntent(nextSurface)
      if (nextSurface === surfaceRef.current) return
      surfaceRef.current = nextSurface
      setDetailsOpen(false)
      if (nextSurface !== 'overview') {
        setIsRoutePending(true)
        void preloadDemo(nextSurface).finally(() => setIsRoutePending(false))
      }
      setSurface(nextSurface)
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

  useEffect(() => {
    preloadDemoOnIntent(surface)
  }, [surface])

  if (isDesignSystemHash()) {
    return (
      <Suspense fallback={<div className="demo-loading" role="status">Opening design system…</div>}>
        <DesignSystemGallery />
      </Suspense>
    )
  }

  if (isDocsHash()) {
    return (
      <Suspense fallback={<div className="demo-loading" role="status">Opening guides…</div>}>
        <DocsApp />
      </Suspense>
    )
  }

  return (
    <div className="app-shell ds" data-surface={surface} data-layout="univer" data-navigation="text">
      <a className="skip-link" href="#main-content">Skip to demo</a>
      <AppHeader sidecar={sidecar} scheme={scheme} onScheme={(next) => { persistColorScheme(next); setScheme(next) }} />
      <div className="app-frame">
      <aside className="app-sidebar">
        <ToolNavigation surface={surface} />
      </aside>
      <main className="app-main" id="main-content" data-workbench-surface={surface}>
        <div className="route-progress" data-active={isRoutePending ? 'true' : undefined} aria-hidden="true"><i /></div>
        <span className="visually-hidden" role="status" aria-live="polite">
          {isRoutePending && requestedTitle ? `Opening ${requestedTitle}…` : ''}
        </span>
        {surface === 'overview' || !demo ? (
          <OverviewPage sidecar={sidecar} />
        ) : (
          <DemoSection
            demo={demo}
            sidecar={sidecar}
            proofOpen={detailsOpen}
            onOpenProof={() => setDetailsOpen(true)}
            proofButtonRef={detailsButtonRef}
          />
        )}
        {demo ? <SourceProofDrawer key={demo.surface} demo={demo} open={detailsOpen} onClose={closeDetails} returnFocusRef={detailsButtonRef} /> : null}
      </main>
      </div>
    </div>
  )
}
