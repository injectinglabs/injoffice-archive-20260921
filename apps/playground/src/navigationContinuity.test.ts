import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const appSource = readFileSync(new URL('./App.tsx', import.meta.url), 'utf8')
const overviewSource = readFileSync(new URL('./pages/OverviewPage.tsx', import.meta.url), 'utf8')
const styles = readFileSync(new URL('./workbench.css', import.meta.url), 'utf8')
const viteConfig = readFileSync(new URL('../vite.config.ts', import.meta.url), 'utf8')

describe('playground navigation continuity', () => {
  it('keeps all section shells and lazily mounts independent live previews', () => {
    expect(appSource).toContain('data-layout="univer"')
    expect(appSource).toContain('SCROLL_SECTIONS.filter(section => section.demo).map(section => <DemoSection')
    expect(appSource).toContain('key={section.key}')
    expect(appSource).toContain('<DemoComponent fixedTool={section.tool} initialHash={initialHash} />')
    expect(appSource).toContain('location.hash !== handledHash.current')
    expect(appSource).toContain('anchorTarget.current = section.key')
    expect(appSource).toContain('new ResizeObserver(onLayout)')
    expect(appSource).toContain("['wheel', 'touchstart', 'pointerdown', 'keydown']")
    expect(appSource).toContain('anchorTarget.current = null')
    expect(appSource).toContain('window.removeEventListener(event, releaseAnchor)')
    expect(appSource).toContain('new IntersectionObserver')
    expect(appSource).toContain('entry.isIntersecting')
    expect(appSource).toContain("data-section-loaded={loadState === 'ready'}")
    expect(appSource).toContain('attempt.current === id')
    expect(appSource).toContain('observer.disconnect()')
    expect(appSource).toContain('Your changes stay here while you explore other sections.')
    expect(appSource).not.toContain('createRouteLoader')
    expect(appSource).not.toContain('isDocsHash')
    expect(appSource).not.toContain('isDesignSystemHash')
    expect(appSource).not.toContain('#/guides')
    expect(appSource).not.toContain('#/design-system')
    expect(appSource).toContain('`${import.meta.env.BASE_URL}logo.svg`')
  })

  it('isolates recoverable chunk and rendering errors to the affected section', () => {
    expect(appSource).toContain("loadState === 'error' ? 'demo-section-error' : 'demo-section-placeholder'")
    expect(appSource).toContain("loadState === 'error' ? 'Retry' : 'Load demo'")
    expect(appSource).toContain('onClick={load}')
    expect(appSource).toContain('Other demos are still available.')
    expect(appSource).toContain('class SectionBoundary extends Component')
    expect(appSource).toContain('getDerivedStateFromError')
    expect(appSource).toContain('<SectionBoundary key={revision}')
    expect(appSource).toContain('Other sections are still available.')
    expect(appSource).not.toContain('preloadDemo(nextSurface).finally')
    const sectionSource = appSource.slice(appSource.indexOf('function DemoSection('), appSource.indexOf('class SectionBoundary'))
    expect(sectionSource).toContain('<Suspense fallback=')
  })

  it('makes walkthroughs discoverable and resets the drawer for each AI format', () => {
    expect(appSource).toContain('Guide &amp; source')
    expect(appSource).toContain('<SourceProofDrawer key={proofSection.key}')
    expect(appSource).toContain('detailsButtonRef.current = button; setProofSection(section)')
    expect(overviewSource).toContain('readShowcasePreferences()')
    expect(overviewSource).toContain('persistShowcasePreferences({ query, task, format })')
  })

  it('clears the desktop sidebar basis when navigation becomes a mobile row', () => {
    expect(styles).toMatch(/@media \(max-width: 760px\)[\s\S]*?\.app-sidebar \{[^}]*flex: 0 0 auto;/)
  })

  it('warms route chunks for pointer and keyboard navigation intent', () => {
    for (const source of [appSource, overviewSource]) {
      expect(source).toContain('onPointerEnter={warmRoute}')
      expect(source).toContain('onPointerDown={warmRoute}')
      expect(source).toContain('onFocus={warmRoute}')
    }
  })

  it('pre-optimizes dependencies imported only by lazy routes in development', () => {
    expect(viteConfig).toContain("entries: ['index.html', 'src/**/*.{ts,tsx}']")
  })

  it('updates the active location passively without stealing focus or adding history entries', () => {
    const updater = appSource.slice(appSource.indexOf('const update = () => {'), appSource.indexOf('const schedule = () =>'))
    expect(updater).toContain('activeSectionKey(')
    expect(updater).toContain("history.replaceState(history.state, '', nextHash)")
    expect(updater).toContain('sectionHashes.current.get(section.key)')
    expect(updater).not.toMatch(/pushState|\.focus\(|scrollIntoView|dispatchEvent|location\.hash\s*=/)
    expect(appSource).toContain("window.addEventListener('scroll', schedule, { passive: true })")
    expect(appSource).toContain('new ResizeObserver(onLayout)')
    expect(appSource).toContain('requestAnimationFrame(update)')
    expect(appSource).toContain("window.removeEventListener('scroll', schedule)")
    expect(appSource).toContain('cancelAnimationFrame(frame)')
  })
})
