import { useEffect, useMemo, useState } from 'react'
import {
  DEMO_FORMATS,
  DEMOS,
  DEMO_TASKS,
  preloadDemoOnIntent,
  type DemoFormat,
  type DemoTask,
} from '../demoRegistry'
import { filterShowcaseItems, showcaseItems } from '../showcaseCatalog'
import { agentHref, surfaceHref, type AgentTool } from '../route'
import { persistShowcasePreferences, readShowcasePreferences } from '../showcasePreferences'

const ALL_TASKS = 'All tasks' as const
const ALL_FORMATS = 'All formats' as const

const AGENT_STARTERS: { tool: AgentTool; format: string; title: string; description: string; file: string }[] = [
  { tool: 'sheets', format: 'XLSX', title: 'Update a workstream status', description: 'Move a workstream forward in a populated launch-readiness plan.', file: 'Launch-readiness plan' },
  { tool: 'docs', format: 'DOCX', title: 'Revise a document title', description: 'Give the Northstar launch brief a new title without rebuilding the document.', file: 'Northstar launch brief' },
  { tool: 'slides', format: 'PPTX', title: 'Update a presentation title', description: 'Rename the opening slide in a real Northstar launch-review deck.', file: 'Northstar launch review' },
  { tool: 'pdf', format: 'PDF', title: 'Rotate a PDF page', description: 'Change one page’s orientation in a four-page operating review.', file: 'Northstar operating review' },
]

function FilterButton<T extends string>({ value, selected, onSelect }: { value: T; selected: boolean; onSelect: (value: T) => void }) {
  return <button type="button" aria-pressed={selected} onClick={() => onSelect(value)}>{value}</button>
}

export default function OverviewPage({ sidecar }: { sidecar: 'checking' | 'connected' | 'offline' }) {
  const [initialFilters] = useState(() => readShowcasePreferences())
  const [query, setQuery] = useState(initialFilters.query)
  const [task, setTask] = useState<DemoTask | typeof ALL_TASKS>(initialFilters.task)
  const [format, setFormat] = useState<DemoFormat | typeof ALL_FORMATS>(initialFilters.format)
  useEffect(() => { persistShowcasePreferences({ query, task, format }) }, [query, task, format])
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

  return (
    <div className="overview-page">
      <section className="task-launch" aria-labelledby="showcase-title">
        <header className="task-launch__intro">
          <h1 id="showcase-title" tabIndex={-1}>Give an agent a document task</h1>
          <p>Pick a sample file and a focused edit. Review what will change, approve it yourself, and download the verified result.</p>
          <p className="task-launch__boundary"><strong>Simulated agent · real document operations</strong><span>A deterministic mock proposes the edit. No LLM, API key, or external AI service.</span></p>
        </header>
        <ul className="task-launch__tasks" aria-label="Choose a document task">
          {AGENT_STARTERS.map((starter) => (
            <li key={starter.tool}>
              <a className="task-launch__task" href={agentHref(starter.tool)} data-agent-task={starter.tool} onPointerEnter={() => preloadDemoOnIntent('agent')} onFocus={() => preloadDemoOnIntent('agent')}>
                <span className="task-launch__format">{starter.format}</span>
                <span className="task-launch__task-copy"><strong>{starter.title}</strong><span>{starter.description}</span><small>{starter.file} · bundled sample</small></span>
                <span className="task-launch__open">Try task</span>
              </a>
            </li>
          ))}
        </ul>
        <ol className="task-launch__workflow" aria-label="How the guided demo works">
          <li><strong>Choose a task</strong><span>Inspect a real sample file.</span></li>
          <li><strong>Review the difference</strong><span>See the proposed change.</span></li>
          <li><strong>Approve the edit</strong><span>You authorize the file write.</span></li>
          <li><strong>Download verified output</strong><span>Reopen and check the result.</span></li>
        </ol>
        <div className="task-launch__footnote"><p>These guided tasks cover specific edits in bundled files, not unrestricted AI. Choosing a task never applies or approves a change.</p><button type="button" onClick={focusCatalog}>Browse all demos</button></div>
      </section>
      <section className="showcase-catalog" id="showcase-catalog" aria-labelledby="showcase-catalog-title">
        <header className="showcase-heading">
          <div>
            <h2 id="showcase-catalog-title">Explore the document engines</h2>
            <p>Scroll through the demos below, or jump to a section from the navigation. Filter sixteen focused surfaces across twenty-six packages to find a starting point. These are bounded proofs, not a claim of unrestricted Microsoft Office parity.</p>
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
                      <b>Jump to demo</b>
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
    </div>
  )
}
