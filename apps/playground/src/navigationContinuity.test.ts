import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const appSource = readFileSync(new URL('./App.tsx', import.meta.url), 'utf8')
const overviewSource = readFileSync(new URL('./pages/OverviewPage.tsx', import.meta.url), 'utf8')
const styles = readFileSync(new URL('./workbench.css', import.meta.url), 'utf8')
const viteConfig = readFileSync(new URL('../vite.config.ts', import.meta.url), 'utf8')

describe('playground navigation continuity', () => {
  it('mounts one Univer-style preview page at a time from the hash route', () => {
    expect(appSource).toContain('data-layout="univer"')
    expect(appSource).toContain("surface === 'overview' || !demo")
    expect(appSource).toContain('<DemoSection')
    expect(appSource).toContain('<DemoComponent key={`${demo.surface}:${agentTool?.tool ?? \'page\'}:${revision}`} />')
    expect(appSource).not.toContain('DEMOS.map((item) => (')
    expect(appSource).not.toContain('className="demo-stage-placeholder"')
    expect(appSource).toContain('createRouteLoader')
    expect(appSource).toContain('startTransition(() => setRoute(next))')
    expect(appSource).toContain('aria-busy={isRoutePending}')
    expect(appSource).toContain('routeLoader.cancel()')
    expect(appSource).not.toContain('isDocsHash')
    expect(appSource).not.toContain('isDesignSystemHash')
    expect(appSource).not.toContain('#/guides')
    expect(appSource).not.toContain('#/design-system')
    expect(appSource).toContain('`${import.meta.env.BASE_URL}logo.svg`')
  })

  it('offers recoverable chunk errors without unmounting the current page', () => {
    expect(appSource).toContain('className="route-error" role="alert"')
    expect(appSource).toContain('routeLoader.load(routeError)')
    expect(appSource).not.toContain('preloadDemo(nextSurface).finally')
    expect(appSource.indexOf('<Suspense fallback=')).toBeLessThan(appSource.indexOf("{surface === 'overview' || !demo"))
  })

  it('makes walkthroughs discoverable and resets the drawer for each AI format', () => {
    expect(appSource).toContain('Guide &amp; source')
    expect(appSource).toContain('key={drawerKey}')
    expect(appSource).toContain('`agent:${parseAgentTool(hash)}`')
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

  it('uses an in-context progress line with a reduced-motion state', () => {
    expect(appSource).toContain('className="route-progress"')
    expect(appSource).toContain('role="status" aria-live="polite"')
    expect(styles).toContain('.route-progress[data-active="true"]')
    expect(styles).toMatch(/@media \(prefers-reduced-motion: reduce\)[\s\S]*\.route-progress\[data-active="true"\] i \{ animation: none; \}/)
  })
})
