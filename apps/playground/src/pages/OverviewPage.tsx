import { DEMO_GROUPS, DEMOS, preloadDemoOnIntent } from '../demoRegistry'
import { surfaceHref } from '../route'

export default function OverviewPage({ sidecar }: { sidecar: 'checking' | 'connected' | 'offline' }) {
  const warmSheets = () => { preloadDemoOnIntent('sheets') }
  const runtimeLabel = sidecar === 'checking'
    ? 'Browser engines ready · checking optional server'
    : sidecar === 'connected'
      ? 'Browser engines ready · server fallback connected'
      : 'Browser engines ready · server fallback offline'

  return (
    <div className="overview-page">
      <section className="overview-hero" aria-labelledby="overview-title">
        <div className="overview-copy">
          <p className="hero-note"><span aria-hidden="true">↻</span> Original Office bytes stay authoritative</p>
          <h1 id="overview-title">Edit Office files without rebuilding what you don’t understand.</h1>
          <p className="hero-lede">Open-source TypeScript, Go, and browser-WASM engines for agents and web applications. Inspect a real file, apply a bounded change, reopen the exact output, and preserve everything outside the edit.</p>
          <div className="hero-actions">
            <a className="primary-action" href={`${surfaceHref('sheets')}?view=native`} onPointerEnter={warmSheets} onPointerDown={warmSheets} onFocus={warmSheets}>Run the native XLSX proof</a>
            <a className="secondary-action" href="#/guides">Read the guides</a>
          </div>
          <dl className="hero-facts">
            <div><dt>24</dt><dd>composable packages</dd></div>
            <div><dt>15</dt><dd>proof surfaces</dd></div>
            <div><dt>4</dt><dd>document formats</dd></div>
          </dl>
        </div>
        <div className="authority-ledger" aria-label="Native XLSX proof sequence">
          <header>
            <div><strong>launch-readiness-plan.xlsx</strong><span>Bundled, repository-owned business fixture</span></div>
            <span className="authority-ledger__runtime"><i aria-hidden="true" />Browser-local</span>
          </header>
          <ol>
            <li><span>Extract</span><div><strong>Read the original package</strong><code>sha256:bdf753af2b…612f0d</code></div></li>
            <li><span>Guard</span><div><strong>Bind one typed change to that revision</strong><code>cell.set_value · Data!A1</code></div></li>
            <li><span>Apply</span><div><strong>Patch only the requested XML</strong><code>Go engine in a browser Worker</code></div></li>
            <li><span>Verify</span><div><strong>Reopen the exact returned bytes</strong><code>value confirmed · revision advanced</code></div></li>
          </ol>
          <footer><span aria-hidden="true">✓</span><div><strong>Fail closed</strong><small>Unsafe or stale changes produce no replacement file.</small></div></footer>
          <a href={`${surfaceHref('sheets')}?view=native`} onPointerEnter={warmSheets} onPointerDown={warmSheets} onFocus={warmSheets}>Open the working proof <span aria-hidden="true">›</span></a>
        </div>
      </section>

      <section className="capability-index" id="tools" aria-labelledby="tools-title">
        <header><div><h2 id="tools-title">Explore the engine labs</h2><p>Fifteen focused surfaces demonstrate the twenty-four packages. Shared infrastructure appears inside the workflows it powers. These are bounded proofs, not a claim of unrestricted Microsoft Office parity.</p></div><span className={`overview-runtime overview-runtime--${sidecar}`} role="status" aria-live="polite">{runtimeLabel}</span></header>
        <div className="capability-groups">
          {DEMO_GROUPS.map((group) => (
            <section key={group}>
              <h3>{group}</h3>
              {DEMOS.filter((demo) => demo.group === group).map((demo) => {
                const warmRoute = () => { preloadDemoOnIntent(demo.surface) }
                return (
                  <a
                    className={`capability-row capability-row--${demo.accent}`}
                    href={surfaceHref(demo.surface)}
                    key={demo.surface}
                    onPointerEnter={warmRoute}
                    onPointerDown={warmRoute}
                    onFocus={warmRoute}
                  >
                    <span className="capability-row-glyph" aria-hidden="true">{demo.glyph}</span>
                    <span><strong>{demo.navTitle}</strong><small>{demo.description}</small></span>
                    <em>{demo.runtime}</em>
                  </a>
                )
              })}
            </section>
          ))}
        </div>
      </section>
    </div>
  )
}
