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
import type { DemoDefinition } from './demoRegistry'
import { WORKSPACE_DEMOS, preloadWorkspace, preloadWorkspaceOnIntent, workspaceProofDemo } from './workspaceRegistry'
import { resolveToolWorkspace } from './toolWorkspaces'
import { surfaceHref, type Surface } from './route'
import { SCROLL_SECTIONS, sectionForHash, workspaceNavigationHash, activeSectionKey, type ScrollSection } from './scrollSections'
import { createDemoRetention } from './demoRetention'

function isModifiedClick(event: MouseEvent<HTMLAnchorElement>) {
  return event.metaKey || event.ctrlKey || event.shiftKey || event.altKey || event.button !== 0
}

function ToolNavigation({ surface, remembered }: { surface: Surface; remembered: Map<string, string> }) {
  const onClick = (event: MouseEvent<HTMLAnchorElement>) => {
    if (isModifiedClick(event)) return
    const href = event.currentTarget.getAttribute('href')
    if (href && sectionForHash(location.hash).href === href) {
      event.preventDefault()
      window.dispatchEvent(new HashChangeEvent('hashchange'))
    } else if (href) {
      const previous = remembered.get(sectionForHash(href).key)
      if (previous) { event.preventDefault(); window.location.hash = previous }
    }
  }
  return <nav className="tool-nav" id="demo-navigation" aria-label="InjOffice tools">
    {WORKSPACE_DEMOS.map(demo => <a key={demo.surface} href={surfaceHref(demo.surface)}
      aria-current={surface === demo.surface ? 'location' : undefined}
      onPointerEnter={() => preloadWorkspaceOnIntent(demo.surface)}
      onPointerDown={() => preloadWorkspaceOnIntent(demo.surface)}
      onFocus={() => preloadWorkspaceOnIntent(demo.surface)}
      onClick={onClick}>{demo.navTitle}</a>)}
  </nav>
}

function ColorSchemeToggle({ scheme, onScheme }: { scheme: ColorScheme; onScheme: (next: ColorScheme) => void }) {
  return (
    <div className="scheme-toggle" role="group" aria-label="Color scheme">
      <button type="button" aria-pressed={scheme === 'light'} onClick={() => onScheme('light')}>Light</button>
      <button type="button" aria-pressed={scheme === 'dark'} onClick={() => onScheme('dark')}>Dark</button>
    </div>
  )
}

function AppHeader({ scheme, onScheme }: { scheme: ColorScheme; onScheme: (next: ColorScheme) => void }) {
  return (
    <header className="app-header">
      <div className="app-header-inner">
        <a className="app-brand" href={surfaceHref('sheets')} aria-label="InjOffice demo start" onClick={event => {
          if (!isModifiedClick(event) && location.hash === surfaceHref('sheets')) {
            event.preventDefault()
            window.dispatchEvent(new HashChangeEvent('hashchange'))
          }
        }}>
          <img className="app-logo" src={`${import.meta.env.BASE_URL}logo.svg`} alt="" width={32} height={32} />
          <span><strong>InjOffice</strong></span>
        </a>
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
  proofOpen,
  onOpenProof,
  section,
  requested,
  requestVersion,
  initialHash,
}: {
  demo: DemoDefinition
  proofOpen: boolean
  onOpenProof: (button: HTMLButtonElement) => void
  section: ScrollSection
  requested: boolean
  requestVersion: number
  initialHash: string
}) {
  const [revision, setRevision] = useState(0)
  const [loadState, setLoadState] = useState<'idle' | 'loading' | 'ready' | 'error'>('idle')
  const sectionRef = useRef<HTMLElement>(null)
  const attempt = useRef(0)
  const loading = useRef(false)
  const touched = useRef(false)
  const explicitlyClosed = useRef(false)
  const [retention] = useState(() => createDemoRetention(() => {
    attempt.current++
    loading.current = false
    setLoadState('idle')
  }))
  const touch = () => { touched.current = true; retention.touch() }
  const busy = () => Boolean(sectionRef.current?.querySelector('[data-demo-busy="true"]'))
  const resetDemo = () => {
    if (busy()) { window.alert('Wait for the current operation to finish before resetting this demo.'); return }
    if (touched.current && !window.confirm('Reset this demo? Your current edits and progress will be lost. Download any files you want to keep first.')) return
    touched.current = false
    retention.reset()
    setRevision(value => value + 1)
  }
  const closeDemo = () => {
    if (busy()) { window.alert('Wait for the current operation to finish before closing this demo.'); return }
    if (touched.current && !window.confirm('Close this demo? Your current edits and progress will be lost. Download any files you want to keep first.')) return
    explicitlyClosed.current = true
    touched.current = false
    retention.setLoaded(false)
    retention.reset()
    attempt.current++
    loading.current = false
    setLoadState('idle')
  }
  const load = useCallback(() => {
    if (loading.current) return
    explicitlyClosed.current = false
    loading.current = true
    const id = ++attempt.current
    setLoadState('loading')
    void preloadWorkspace(demo.surface).then(() => {
      if (attempt.current === id) setLoadState('ready')
    }, () => {
      if (attempt.current === id) { loading.current = false; setLoadState('error') }
    })
  }, [demo.surface])
  useEffect(() => () => { attempt.current++; loading.current = false }, [])
  useEffect(() => () => retention.dispose(), [retention])
  useEffect(() => { retention.setLoaded(loadState === 'ready') }, [loadState, retention])
  useEffect(() => { if (requested) load() }, [requested, requestVersion, load])
  useEffect(() => {
    const element = sectionRef.current
    if (!element) return
    const observer = new IntersectionObserver(entries => {
      const visible = entries.some(entry => entry.isIntersecting)
      retention.setVisible(visible)
      if (visible && loadState === 'idle' && !explicitlyClosed.current) load()
    }, { rootMargin: '160px 0px' })
    observer.observe(element)
    return () => observer.disconnect()
  }, [load, loadState, retention])
  const DemoComponent = demo.component as ComponentType<{ initialHash?: string }>
  const title = demo.title
  const description = demo.description
  const Heading = demo.surface === 'sheets' ? 'h1' : 'h2'
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
          <div className="demo-title-line">
            <Heading id={`demo-title-${section.key}`} tabIndex={-1}>{title}</Heading>
          </div>
          <p>{description}</p>
        </div>
        <div className="demo-context-actions">
          <details className="demo-options">
          <summary>Options</summary>
          <div>
          <button
            className="demo-reset-trigger"
            type="button"
            disabled={loadState !== 'ready'}
            onClick={resetDemo}
          >
            Reset demo
          </button>
          {loadState === 'ready' && <button className="demo-reset-trigger" type="button" onClick={closeDemo}>Close demo</button>}
          </div>
          </details>
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
        </div>
      </header>
      <section className="demo-preview" aria-label={`${title} preview`}>
        <div className="demo-stage" data-accent={demo.accent} aria-label={`${title} interactive demo`} onPointerDownCapture={touch} onKeyDownCapture={touch} onInputCapture={touch} onClickCapture={touch}>
          {loadState === 'ready' ? <SectionBoundary key={revision} onRetry={() => setRevision(value => value + 1)}>
            <Suspense fallback={<div className="demo-loading" role="status">Opening {title}…</div>}>
              <DemoComponent initialHash={initialHash} />
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
  const [route, setRoute] = useState(() => ({ surface: sectionForHash(location.hash).surface, hash: workspaceNavigationHash(location.hash) }))
  const { surface, hash } = route
  const [scheme, setScheme] = useState<ColorScheme>(() => currentColorScheme())
  const [proofSection, setProofSection] = useState<(ScrollSection & { featureHash: string }) | null>(null)
  const [detailsOpen, setDetailsOpen] = useState(false)
  const [requestedKey, setRequestedKey] = useState(() => sectionForHash(location.hash).key)
  const [requestVersion, setRequestVersion] = useState(0)
  const detailsButtonRef = useRef<HTMLButtonElement>(null)
  const closeDetails = useCallback(() => setDetailsOpen(false), [])
  const sectionHashes = useRef(new Map<string, string>())
  const navigating = useRef(false)
  const anchorTarget = useRef<string | null>(null)
  const anchorViewTop = useRef<number | null>(null)
  const viewIntent = useRef<{ key: string; top: number } | null>(null)
  const handledHash = useRef(workspaceNavigationHash(location.hash))
  const pinAnchor = useCallback(() => {
    if (!anchorTarget.current) return
    const element = document.getElementById(`demo-${anchorTarget.current}`)
    if (anchorViewTop.current !== null) {
      const controls = element?.querySelector('.tool-workspace__navigation')
      if (controls) {
        const delta = controls.getBoundingClientRect().top - anchorViewTop.current
        if (Math.abs(delta) > 1) window.scrollBy({ top: delta, behavior: 'instant' })
      }
    } else element?.scrollIntoView({ block: 'start', behavior: 'instant' })
  }, [])

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
    const internalView = (event: Event) => {
      const tool = (event as CustomEvent<{ tool?: string }>).detail?.tool
      const section = SCROLL_SECTIONS.find(item => item.key === tool)
      const controls = section && document.getElementById(`demo-${section.key}`)?.querySelector('.tool-workspace__navigation')
      if (section && controls) viewIntent.current = { key: section.key, top: controls.getBoundingClientRect().top }
    }
    const navigate = (focus = true) => {
      const destination = workspaceNavigationHash(location.hash)
      if (location.hash !== destination) history.replaceState(history.state, '', destination)
      const section = sectionForHash(location.hash)
      const intent = viewIntent.current
      viewIntent.current = null
      const changingView = focus && intent?.key === section.key
      anchorTarget.current = section.key
      anchorViewTop.current = changingView ? intent.top : null
      handledHash.current = location.hash || section.href
      sectionHashes.current.set(section.key, location.hash || section.href)
      setRequestedKey(section.key)
      setRequestVersion(value => value + 1)
      setRoute({ surface: section.surface, hash: location.hash || section.href })
      setDetailsOpen(false)
      navigating.current = true
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(() => {
        const element = document.getElementById(`demo-${section.key}`)
        pinAnchor()
        if (focus && !changingView) element?.querySelector<HTMLElement>('h1, h2')?.focus({ preventScroll: true })
        navigating.current = false
      })
    }
    const sync = () => navigate()
    window.addEventListener('injoffice:workspace-view', internalView)
    window.addEventListener('hashchange', sync)
    navigate(false)
    return () => {
      window.removeEventListener('hashchange', sync)
      window.removeEventListener('injoffice:workspace-view', internalView)
      cancelAnimationFrame(frame)
    }
  }, [pinAnchor])

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
      const atEnd = window.scrollY + window.innerHeight >= document.documentElement.scrollHeight - 2
      const key = anchorTarget.current ?? activeSectionKey(SCROLL_SECTIONS.map(section => ({ key: section.key, top: document.getElementById(`demo-${section.key}`)?.getBoundingClientRect().top ?? Infinity })), top + 32, atEnd)
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
    // Lazy sections above a clicked destination can grow after its first
    // scroll. Preserve that destination until the user takes over scrolling.
    const onLayout = () => {
      if (anchorTarget.current && document.body.style.overflow !== 'hidden') {
        pinAnchor()
      }
      schedule()
    }
    const releaseAnchor = () => { anchorTarget.current = null; anchorViewTop.current = null }
    const intentEvents = ['wheel', 'touchstart', 'pointerdown', 'keydown'] as const
    intentEvents.forEach(event => window.addEventListener(event, releaseAnchor, { passive: true }))
    const observer = new ResizeObserver(onLayout)
    const main = document.querySelector('.app-main')
    if (main) observer.observe(main)
    return () => {
      cancelAnimationFrame(frame)
      window.removeEventListener('scroll', schedule)
      window.removeEventListener('resize', schedule)
      observer.disconnect()
      intentEvents.forEach(event => window.removeEventListener(event, releaseAnchor))
    }
  }, [pinAnchor])

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
    preloadWorkspaceOnIntent(surface)
  }, [surface])

  return (
    <div className="app-shell ds" data-surface={surface} data-layout="univer" data-navigation="scroll">
      <a className="skip-link" href="#main-content" onClick={event => {
        event.preventDefault()
        document.getElementById('main-content')?.focus({ preventScroll: true })
      }}>Skip to demo</a>
      <AppHeader scheme={scheme} onScheme={(next) => { persistColorScheme(next); setScheme(next) }} />
      <div className="app-frame">
      <aside className="app-sidebar">
        <ToolNavigation surface={surface} remembered={sectionHashes.current} />
      </aside>
      <main className="app-main" id="main-content" tabIndex={-1} data-workbench-surface={surface} aria-busy={false}>
        {SCROLL_SECTIONS.filter(section => section.demo).map(section => <DemoSection
            key={section.key}
            section={section}
            demo={section.demo!}
            requested={requestedKey === section.key}
            requestVersion={requestVersion}
            initialHash={sectionHashes.current.get(section.key) ?? section.href}
            proofOpen={detailsOpen && proofSection?.key === section.key}
            onOpenProof={button => { detailsButtonRef.current = button; setProofSection({ ...section, featureHash: sectionHashes.current.get(section.key) ?? section.href }); setDetailsOpen(true) }}
          />)}
        {proofSection?.demo ? <SourceProofDrawer key={`${proofSection.key}:${resolveToolWorkspace(proofSection.featureHash)?.feature}`} demo={workspaceProofDemo(proofSection.featureHash)} open={detailsOpen} onClose={closeDetails} returnFocusRef={detailsButtonRef} /> : null}
      </main>
      </div>
    </div>
  )
}
