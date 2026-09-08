import { Component, Suspense, useCallback, useEffect, useRef, useState, type ComponentType, type MouseEvent, type ReactNode, type RefObject } from 'react'
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
import { DEMO_GROUPS, DEMOS, preloadDemo, preloadDemoOnIntent, type DemoDefinition } from './demoRegistry'
import OverviewPage from './pages/OverviewPage'
import { agentHref, parseAgentTool, parseSurface, surfaceHref, AGENT_TOOLS, type Surface } from './route'
import { SCROLL_SECTIONS, sectionForHash, activeSectionKey, type ScrollSection } from './scrollSections'
import type { AgentTool } from './route'

type SidecarState = 'checking' | 'connected' | 'offline'

const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').replace(/\/$/, '')

function isModifiedClick(event: MouseEvent<HTMLAnchorElement>) {
  return event.metaKey || event.ctrlKey || event.shiftKey || event.altKey || event.button !== 0
}

function ToolNavigation({ surface, hash }: { surface: Surface; hash: string }) {
  const onClick = (event: MouseEvent<HTMLAnchorElement>) => {
    if (isModifiedClick(event)) return
    const href = event.currentTarget.getAttribute('href')
    if (href && location.hash === href) {
      event.preventDefault()
      window.dispatchEvent(new HashChangeEvent('hashchange'))
    }
  }
  const agentTool = surface === 'agent' ? parseAgentTool(hash) : undefined
  return (
    <nav className="tool-nav" id="demo-navigation" aria-label="InjOffice tools">
      <a className="tool-nav-home" href={surfaceHref('overview')} aria-current={surface === 'overview' ? 'location' : undefined} onClick={onClick}>Overview</a>
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
                  aria-current={agentTool === item.tool ? 'location' : undefined}
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
                aria-current={surface === demo.surface ? 'location' : undefined}
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
  section,
  requested,
  initialHash,
}: {
  demo: DemoDefinition
  sidecar: SidecarState
  proofOpen: boolean
  onOpenProof: (button: HTMLButtonElement) => void
  section: ScrollSection
  requested: boolean
  initialHash: string
}) {
  const [revision, setRevision] = useState(0)
  const [loadState, setLoadState] = useState<'idle' | 'loading' | 'ready' | 'error'>('idle')
  const sectionRef = useRef<HTMLElement>(null)
  const attempt = useRef(0)
  const loading = useRef(false)
  const load = useCallback(() => {
    if (loading.current) return
    loading.current = true
    const id = ++attempt.current
    setLoadState('loading')
    void preloadDemo(demo.surface).then(() => {
      if (attempt.current === id) setLoadState('ready')
    }, () => {
      if (attempt.current === id) { loading.current = false; setLoadState('error') }
    })
  }, [demo.surface])
  useEffect(() => () => { attempt.current++; loading.current = false }, [])
  useEffect(() => { if (requested) load() }, [requested, load])
  useEffect(() => {
    const element = sectionRef.current
    if (!element || loadState !== 'idle') return
    const observer = new IntersectionObserver(entries => {
      if (entries.some(entry => entry.isIntersecting)) load()
    }, { rootMargin: '160px 0px' })
    observer.observe(element)
    return () => observer.disconnect()
  }, [load, loadState])
  const DemoComponent = demo.component as ComponentType<{ fixedTool?: AgentTool; initialHash?: string }>
  const agentTool = section.tool ? AGENT_TOOLS.find((item) => item.tool === section.tool) : undefined
  const title = agentTool?.title ?? demo.title
  const description = agentTool?.description ?? demo.description
  return (
    <article
      className="demo-section"
      ref={sectionRef}
      id={`demo-${section.key}`}
      data-scroll-section={section.key}
      data-scroll-state={loadState}
      data-section-loaded={loadState === 'ready'}
      data-accent={demo.accent}
      aria-labelledby={`demo-title-${section.key}`}
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
            <h2 id={`demo-title-${section.key}`} tabIndex={-1}>{title}</h2>
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
            className="source-proof-trigger"
            type="button"
            aria-haspopup="dialog"
            aria-expanded={proofOpen}
            aria-controls="source-proof-drawer"
            onClick={event => onOpenProof(event.currentTarget)}
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
          {loadState === 'ready' ? <SectionBoundary key={revision} onRetry={() => setRevision(value => value + 1)}>
            <Suspense fallback={<div className="demo-loading" role="status">Opening {title}…</div>}>
              <DemoComponent fixedTool={section.tool} initialHash={initialHash} />
            </Suspense>
          </SectionBoundary> : <div className={loadState === 'error' ? 'demo-section-error' : 'demo-section-placeholder'} role={loadState === 'error' ? 'alert' : undefined}>
            <p>{loadState === 'error' ? `Could not load ${title}. Other demos are still available.` : loadState === 'loading' ? `Opening ${title}…` : 'This live demo loads as you reach it. Your changes stay here while you explore other sections.'}</p>
            {loadState !== 'loading' && <button type="button" onClick={load}>{loadState === 'error' ? 'Retry' : 'Load demo'}</button>}
            {loadState === 'error' && <>
              <p>If retrying does not help, reload to fetch a fresh copy of the demo.</p>
              <button type="button" onClick={() => { if (window.confirm('Reload this page? Unsaved demo changes will be lost.')) location.reload() }}>Reload page</button>
            </>}
          </div>}
        </div>
      </section>
    </article>
  )
}

class SectionBoundary extends Component<{ children: ReactNode; onRetry: () => void }, { failed: boolean }> {
  state = { failed: false }
  static getDerivedStateFromError() { return { failed: true } }
  render() {
    return this.state.failed ? <div className="demo-section-error" role="alert"><p>This demo could not start. Other sections are still available.</p><button type="button" onClick={this.props.onRetry}>Retry</button></div> : this.props.children
  }
}

export default function App() {
  const [route, setRoute] = useState(() => ({ surface: parseSurface(), hash: location.hash || surfaceHref('overview') }))
  const { surface, hash } = route
  const [sidecar, setSidecar] = useState<SidecarState>('checking')
  const [scheme, setScheme] = useState<ColorScheme>(() => currentColorScheme())
  const [proofSection, setProofSection] = useState<ScrollSection | null>(null)
  const [detailsOpen, setDetailsOpen] = useState(false)
  const [requestedKey, setRequestedKey] = useState(() => sectionForHash(location.hash).key)
  const detailsButtonRef = useRef<HTMLButtonElement>(null)
  const closeDetails = useCallback(() => setDetailsOpen(false), [])
  const sectionHashes = useRef(new Map<string, string>())
  const navigating = useRef(false)
  const handledHash = useRef(location.hash || surfaceHref('overview'))

  useEffect(() => {
    // Scroll the rail only; scrollIntoView here would move the document too.
    const link = document.querySelector<HTMLElement>('.tool-nav a[aria-current="location"]')
    const rail = document.querySelector<HTMLElement>('.app-sidebar')
    if (!link || !rail) return
    const item = link.getBoundingClientRect(), bounds = rail.getBoundingClientRect()
    if (item.top < bounds.top) rail.scrollTop += item.top - bounds.top
    else if (item.bottom > bounds.bottom) rail.scrollTop += item.bottom - bounds.bottom
    if (item.left < bounds.left) rail.scrollLeft += item.left - bounds.left
    else if (item.right > bounds.right) rail.scrollLeft += item.right - bounds.right
  }, [surface, hash])

  useEffect(() => {
    let frame = 0
    const navigate = (focus = true) => {
      const section = sectionForHash(location.hash)
      handledHash.current = location.hash || section.href
      sectionHashes.current.set(section.key, location.hash || section.href)
      setRequestedKey(section.key)
      setRoute({ surface: section.surface, hash: location.hash || section.href })
      setDetailsOpen(false)
      navigating.current = true
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(() => {
        const element = document.getElementById(`demo-${section.key}`)
        element?.scrollIntoView({ block: 'start', behavior: 'instant' })
        if (focus) element?.querySelector<HTMLElement>('h1, h2')?.focus({ preventScroll: true })
        navigating.current = false
      })
    }
    const sync = () => navigate()
    window.addEventListener('hashchange', sync)
    if (!location.hash) history.replaceState(history.state, '', surfaceHref('overview'))
    navigate(false)
    return () => {
      window.removeEventListener('hashchange', sync)
      cancelAnimationFrame(frame)
    }
  }, [])

  useEffect(() => {
    let frame = 0
    const update = () => {
      frame = 0
      // A new hash can precede its queued hashchange event. Never let a stale
      // scroll/resize frame overwrite that explicit navigation intent.
      if (navigating.current || location.hash !== handledHash.current || document.body.style.overflow === 'hidden') return
      const header = document.querySelector('.app-header')?.getBoundingClientRect().bottom ?? 48
      const rail = document.querySelector('.app-sidebar')?.getBoundingClientRect()
      const top = window.innerWidth <= 760 ? Math.max(header, rail?.bottom ?? 0) : header
      const key = activeSectionKey(SCROLL_SECTIONS.map(section => ({ key: section.key, top: document.getElementById(`demo-${section.key}`)?.getBoundingClientRect().top ?? Infinity })), top + 32)
      const section = SCROLL_SECTIONS.find(item => item.key === key)
      if (!section) return
      const nextHash = sectionHashes.current.get(section.key) ?? section.href
      if (location.hash !== nextHash) history.replaceState(history.state, '', nextHash)
      handledHash.current = nextHash
      setRoute(previous => previous.hash === nextHash ? previous : { surface: section.surface, hash: nextHash })
    }
    const schedule = () => { if (!frame) frame = requestAnimationFrame(update) }
    window.addEventListener('scroll', schedule, { passive: true })
    window.addEventListener('resize', schedule)
    const observer = new ResizeObserver(schedule)
    const main = document.querySelector('.app-main')
    if (main) observer.observe(main)
    return () => {
      cancelAnimationFrame(frame)
      window.removeEventListener('scroll', schedule)
      window.removeEventListener('resize', schedule)
      observer.disconnect()
    }
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

  return (
    <div className="app-shell ds" data-surface={surface} data-layout="univer" data-navigation="scroll">
      <a className="skip-link" href="#main-content" onClick={event => {
        event.preventDefault()
        document.getElementById('main-content')?.focus({ preventScroll: true })
      }}>Skip to demo</a>
      <AppHeader sidecar={sidecar} scheme={scheme} onScheme={(next) => { persistColorScheme(next); setScheme(next) }} />
      <div className="app-frame">
      <aside className="app-sidebar">
        <ToolNavigation surface={surface} hash={hash} />
      </aside>
      <main className="app-main" id="main-content" tabIndex={-1} data-workbench-surface={surface} aria-busy={false}>
        <section className="demo-section demo-section--overview" id="demo-overview" data-scroll-section="overview" data-scroll-state="ready" aria-labelledby="showcase-title">
          <OverviewPage sidecar={sidecar} />
        </section>
        {SCROLL_SECTIONS.filter(section => section.demo).map(section => <DemoSection
            key={section.key}
            section={section}
            demo={section.demo!}
            sidecar={sidecar}
            requested={requestedKey === section.key}
            initialHash={sectionHashes.current.get(section.key) ?? section.href}
            proofOpen={detailsOpen && proofSection?.key === section.key}
            onOpenProof={button => { detailsButtonRef.current = button; setProofSection(section); setDetailsOpen(true) }}
          />)}
        {proofSection?.demo ? <SourceProofDrawer key={proofSection.key} demo={proofSection.demo} open={detailsOpen} onClose={closeDetails} returnFocusRef={detailsButtonRef} /> : null}
      </main>
      </div>
    </div>
  )
}
