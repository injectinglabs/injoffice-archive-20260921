import { useMemo, useState } from 'react'
import {
  DEMO_FORMATS,
  DEMOS,
  DEMO_TASKS,
  preloadDemoOnIntent,
  type DemoFormat,
  type DemoTask,
} from '../demoRegistry'
import { filterShowcaseItems, showcaseItems } from '../showcaseCatalog'
import { surfaceHref } from '../route'

const ALL_TASKS = 'All tasks' as const
const ALL_FORMATS = 'All formats' as const

function FilterButton<T extends string>({ value, selected, onSelect }: { value: T; selected: boolean; onSelect: (value: T) => void }) {
  return <button type="button" aria-pressed={selected} onClick={() => onSelect(value)}>{value}</button>
}

export default function OverviewPage({ sidecar }: { sidecar: 'checking' | 'connected' | 'offline' }) {
  const [query, setQuery] = useState('')
  const [task, setTask] = useState<DemoTask | typeof ALL_TASKS>(ALL_TASKS)
  const [format, setFormat] = useState<DemoFormat | typeof ALL_FORMATS>(ALL_FORMATS)
  const warmSheets = () => { preloadDemoOnIntent('sheets') }
  const runtimeLabel = sidecar === 'checking'
    ? 'Browser engines ready · checking optional server'
    : sidecar === 'connected'
      ? 'Browser engines ready · server fallback connected'
      : 'Browser engines ready · server fallback offline'
  const catalog = useMemo(() => showcaseItems(DEMOS), [])
  const matches = useMemo(() => filterShowcaseItems(catalog, { query, task, format }), [catalog, query, task, format])
  const hasFilters = query.length > 0 || task !== ALL_TASKS || format !== ALL_FORMATS
  const resetFilters = () => {
    setQuery('')
    setTask(ALL_TASKS)
    setFormat(ALL_FORMATS)
  }
  const focusCatalog = () => {
    document.querySelector('#showcase-query')?.scrollIntoView({ block: 'center' })
    document.querySelector<HTMLInputElement>('#showcase-query')?.focus({ preventScroll: true })
  }

  const authorityProof = (
      <section className="overview-hero" aria-labelledby="overview-title">
        <div className="overview-copy">
          <p className="hero-note"><span aria-hidden="true">↻</span> Original Office bytes stay authoritative</p>
          <h2 id="overview-title">Choose the job. Keep the source file.</h2>
          <p className="hero-lede">Open-source TypeScript, Go, and browser-WASM engines for agents and web applications. Inspect a real file, apply a bounded change, reopen the exact output, and preserve everything outside the edit.</p>
          <div className="hero-actions">
            <button className="primary-action" type="button" onClick={focusCatalog}>Find a working proof</button>
            <a className="secondary-action" href={`${surfaceHref('sheets')}?view=native`} onPointerEnter={warmSheets} onPointerDown={warmSheets} onFocus={warmSheets}>Run the native XLSX proof</a>
          </div>
          <dl className="hero-facts">
            <div><dt>26</dt><dd>composable packages</dd></div>
            <div><dt>16</dt><dd>proof surfaces</dd></div>
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
        </div>
      </section>
  )

  return (
    <div className="overview-page">
      <section className="showcase-catalog" id="showcase-catalog" aria-labelledby="showcase-title">
        <header className="showcase-heading">
          <div>
            <h1 id="showcase-title" tabIndex={-1}>Start with what you need to do</h1>
            <p>Filter sixteen focused surfaces across the twenty-six packages. Shared infrastructure appears inside the workflows it powers. These are bounded proofs, not a claim of unrestricted Microsoft Office parity.</p>
          </div>
          <span className={`overview-runtime overview-runtime--${sidecar}`} role="status" aria-live="polite">{runtimeLabel}</span>
        </header>
        <div className="showcase-shortcuts">
          <a href={`${surfaceHref('sheets')}?view=editor`} onPointerEnter={warmSheets} onPointerDown={warmSheets} onFocus={warmSheets}>Try the spreadsheet editor</a>
          <a href={`${surfaceHref('sheets')}?view=native`} onPointerEnter={warmSheets} onPointerDown={warmSheets} onFocus={warmSheets}>Verify an XLSX round trip</a>
          <a href={`${surfaceHref('sheets')}?view=tools`} onPointerEnter={warmSheets} onPointerDown={warmSheets} onFocus={warmSheets}>Explore workbook tools</a>
        </div>

        <form className="showcase-filters" role="search" onSubmit={(event) => event.preventDefault()}>
          <div className="showcase-search">
            <label htmlFor="showcase-query">Search the working proofs</label>
            <div>
              <span aria-hidden="true">⌕</span>
              <input
                id="showcase-query"
                type="search"
                value={query}
                onChange={(event) => setQuery(event.target.value)}
                placeholder="Try “edit PPTX”, “redact PDF”, or “chart”"
                autoComplete="off"
              />
              {query && <button type="button" onClick={() => setQuery('')} aria-label="Clear search">Clear</button>}
            </div>
          </div>
          <fieldset className="showcase-filter-group">
            <legend>Task</legend>
            <div>
              <FilterButton value={ALL_TASKS} selected={task === ALL_TASKS} onSelect={setTask} />
              {DEMO_TASKS.map((value) => <FilterButton key={value} value={value} selected={task === value} onSelect={setTask} />)}
            </div>
          </fieldset>
          <fieldset className="showcase-filter-group showcase-filter-group--formats">
            <legend>File type</legend>
            <div>
              <FilterButton value={ALL_FORMATS} selected={format === ALL_FORMATS} onSelect={setFormat} />
              {DEMO_FORMATS.map((value) => <FilterButton key={value} value={value} selected={format === value} onSelect={setFormat} />)}
            </div>
          </fieldset>
        </form>

        <div className="showcase-results-heading">
          <p aria-live="polite"><strong>{matches.length}</strong> {matches.length === 1 ? 'proof' : 'proofs'} available</p>
          {hasFilters && <button type="button" onClick={resetFilters}>Reset filters</button>}
        </div>

        {matches.length > 0 ? (
          <ul className="showcase-results" aria-label="Matching working proofs">
            {matches.map((item) => {
              const warmRoute = () => { preloadDemoOnIntent(item.surface) }
              return (
                <li key={item.key}>
                  <a
                    className={`showcase-item showcase-item--${item.accent}`}
                    href={item.href}
                    onPointerEnter={warmRoute}
                    onPointerDown={warmRoute}
                    onFocus={warmRoute}
                  >
                    <span className="showcase-item__glyph" aria-hidden="true">{item.glyph}</span>
                    <span className="showcase-item__body">
                      <span className="showcase-item__title"><strong>{item.title}</strong></span>
                      <span className="showcase-item__description">{item.description}</span>
                      <span className="showcase-item__tasks">{item.subtitle}</span>
                    </span>
                    <span className="showcase-item__meta">
                      <span>{item.formats.join(' + ')}</span>
                      <small>{item.runtime} · {item.minutes} min</small>
                      <b>Open proof</b>
                    </span>
                  </a>
                </li>
              )
            })}
          </ul>
        ) : (
          <div className="showcase-empty" role="status">
            <span aria-hidden="true">⌕</span>
            <div><strong>No working proof matches those filters.</strong><p>Try a broader task, another file type, or fewer search terms.</p></div>
            <button type="button" onClick={resetFilters}>Show all proofs</button>
          </div>
        )}
      </section>
      {authorityProof}
    </div>
  )
}
