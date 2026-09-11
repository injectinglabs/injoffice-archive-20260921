import { Component, Suspense, useCallback, useEffect, useRef, useState, type ComponentType, type MouseEvent, type ReactNode } from 'react'
import {
  applyColorScheme,
  currentColorScheme,
  persistColorScheme,
  preferredColorScheme,
  readStoredColorScheme,
  type ColorScheme,
} from './colorScheme'
import type { DemoDefinition } from './demoRegistry'
import { preloadWorkspace, preloadWorkspaceOnIntent } from './workspaceRegistry'
import { resolveToolWorkspace, TOOL_WORKSPACES, workspaceExamples, workspaceHref } from './toolWorkspaces'
import { surfaceHref, type Surface } from './route'
import { SCROLL_SECTIONS, sectionForHash, workspaceNavigationHash, activeSectionKey, type ScrollSection } from './scrollSections'
import { createDemoRetention } from './demoRetention'

function isModifiedClick(event: MouseEvent<HTMLAnchorElement>) {
  return event.metaKey || event.ctrlKey || event.shiftKey || event.altKey || event.button !== 0
}

function ToolNavigation({ surface, hash, remembered }: { surface: Surface; hash: string; remembered: Map<string, string> }) {
  const onClick = (event: MouseEvent<HTMLAnchorElement>) => {
    if (isModifiedClick(event)) return
    const href = event.currentTarget.getAttribute('href')
    if (!href) return
    const destination = !event.currentTarget.dataset.workspaceFeature
      ? remembered.get(sectionForHash(href).key) ?? href : href
    event.preventDefault()
    if (location.hash === destination) window.dispatchEvent(new HashChangeEvent('hashchange'))
    else window.location.hash = destination
  }
  return <nav className="tool-nav" id="demo-navigation" tabIndex={-1} aria-label="InjOffice tools">
    {TOOL_WORKSPACES.map(workspace => <section className="tool-nav-examples" key={workspace.tool} aria-label={`${workspace.title} examples`}>
      <a className="tool-nav-title" href={surfaceHref(workspace.tool)}
      aria-current={surface === workspace.tool ? 'location' : undefined}
      onPointerEnter={() => preloadWorkspaceOnIntent(workspace.tool)}
      onPointerDown={() => preloadWorkspaceOnIntent(workspace.tool)}
      onFocus={() => preloadWorkspaceOnIntent(workspace.tool)}
      onClick={onClick}>{workspace.title}</a>
      <ul>{workspaceExamples(workspace).map(feature => <li key={feature.id}>
        <a href={workspaceHref(workspace.tool, feature.id)} data-workspace-feature={feature.id}
          aria-current={surface === workspace.tool && resolveToolWorkspace(hash)?.feature === feature.id ? 'true' : undefined}
          onClick={onClick}>{feature.label}</a>
      </li>)}</ul>
    </section>)}
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
          <a className="examples-index-link" href="#demo-navigation" onClick={event => {
            event.preventDefault()
            const navigation = document.getElementById('demo-navigation')
            navigation?.scrollIntoView({ block: 'start' })
            navigation?.focus({ preventScroll: true })
          }}>Examples</a>
          <ColorSchemeToggle scheme={scheme} onScheme={onScheme} />
          <a className="github-link" href="https://github.com/injectinglabs/injoffice" target="_blank" rel="noreferrer">GitHub<span aria-hidden="true">↗</span></a>
        </div>
      </div>
    </header>
  )
}

function DemoSection({
  demo,
  section,
  requested,
  requestVersion,
  initialHash,
}: {
  demo: DemoDefinition
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
  const [retention] = useState(() => createDemoRetention(() => {
    attempt.current++
    loading.current = false
    setLoadState('idle')
  }))
  const touch = () => retention.touch()
  const load = useCallback(() => {
    if (loading.current) return
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
      if (visible && loadState === 'idle') load()
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
  const [requestedKey, setRequestedKey] = useState(() => sectionForHash(location.hash).key)
  const [requestVersion, setRequestVersion] = useState(0)
  const sectionHashes = useRef(new Map<string, string>())
  const navigating = useRef(false)
  const anchorTarget = useRef<string | null>(null)
  const handledHash = useRef(workspaceNavigationHash(location.hash))
  const pinAnchor = useCallback(() => {
    if (!anchorTarget.current) return
    const element = document.getElementById(`demo-${anchorTarget.current}`)
    const target = resolveToolWorkspace(sectionHashes.current.get(anchorTarget.current) ?? '')
    const example = target && document.getElementById(`example-${target.tool}-${target.feature}`)
    const destination = example || element
    destination?.scrollIntoView({ block: 'start', behavior: 'instant' })
  }, [])

  useEffect(() => {
    // Scroll the rail only; scrollIntoView here would move the document too.
    const link = document.querySelector<HTMLElement>('.tool-nav a[aria-current="true"]') ?? document.querySelector<HTMLElement>('.tool-nav a[aria-current="location"]')
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
      const destination = workspaceNavigationHash(location.hash)
      if (location.hash !== destination) history.replaceState(history.state, '', destination)
      const section = sectionForHash(location.hash)
      anchorTarget.current = section.key
      handledHash.current = location.hash || section.href
      sectionHashes.current.set(section.key, location.hash || section.href)
      setRequestedKey(section.key)
      setRequestVersion(value => value + 1)
      setRoute({ surface: section.surface, hash: location.hash || section.href })
      navigating.current = true
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(() => {
        const element = document.getElementById(`demo-${section.key}`)
        pinAnchor()
        const target = resolveToolWorkspace(location.hash)
        const example = target && document.getElementById(`example-${target.tool}-${target.feature}`)
        if (focus) (example || element)?.querySelector<HTMLElement>('h1, h2, h3')?.focus({ preventScroll: true })
        navigating.current = false
      })
    }
    const sync = () => navigate()
    window.addEventListener('hashchange', sync)
    navigate(false)
    return () => {
      window.removeEventListener('hashchange', sync)
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
      const top = header
      const atEnd = window.scrollY + window.innerHeight >= document.documentElement.scrollHeight - 2
      const key = anchorTarget.current ?? activeSectionKey(SCROLL_SECTIONS.map(section => ({ key: section.key, top: document.getElementById(`demo-${section.key}`)?.getBoundingClientRect().top ?? Infinity })), top + 32, atEnd)
      const section = SCROLL_SECTIONS.find(item => item.key === key)
      if (!section) return
      const previousHash = sectionHashes.current.get(section.key) ?? section.href
      const panels = Array.from(document.querySelectorAll<HTMLElement>(`[data-scroll-section="${section.key}"] [data-workspace-panel]`))
      const feature = activeSectionKey(panels.map(panel => ({ key: panel.dataset.workspacePanel!, top: panel.getBoundingClientRect().top })), top + 32, atEnd)
      const previousTarget = resolveToolWorkspace(previousHash)
      const nextHash = anchorTarget.current || !feature || previousTarget?.feature === feature
        ? previousHash : workspaceHref(section.key as 'sheets' | 'docs' | 'slides' | 'pdf', feature)
      sectionHashes.current.set(section.key, nextHash)
      if (location.hash !== nextHash) history.replaceState(history.state, '', nextHash)
      handledHash.current = nextHash
      setRoute(previous => previous.hash === nextHash ? previous : { surface: section.surface, hash: nextHash })
      window.dispatchEvent(new Event('injoffice:workspace-scroll'))
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
    const releaseAnchor = () => { anchorTarget.current = null }
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
        <ToolNavigation surface={surface} hash={hash} remembered={sectionHashes.current} />
      </aside>
      <main className="app-main" id="main-content" tabIndex={-1} data-workbench-surface={surface} aria-busy={false}>
        {SCROLL_SECTIONS.filter(section => section.demo).map(section => <DemoSection
            key={section.key}
            section={section}
            demo={section.demo!}
            requested={requestedKey === section.key}
            requestVersion={requestVersion}
            initialHash={sectionHashes.current.get(section.key) ?? section.href}
          />)}
      </main>
      </div>
    </div>
  )
}
