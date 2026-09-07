import { useMemo, useState, type ReactNode } from 'react'
import {
  DsAvatar,
  DsButton,
  DsCallout,
  DsChip,
  DsField,
  DsInput,
  DsMark,
  DsSegment,
  DsSelect,
  DsTextarea,
} from './primitives'
import './surfaces.css'

function Frame({ children }: { children: ReactNode }) {
  return <div className="ds-sheet-frame">{children}</div>
}

function Head({ pkg, title, runtime }: { pkg: string; title: string; runtime: string }) {
  return (
    <header className="ds-surf-head">
      <div>
        <p className="ds-surf-pkg">{pkg}</p>
        <h3>{title}</h3>
      </div>
      <DsChip tone="blue">{runtime}</DsChip>
    </header>
  )
}

function Proof({ rows }: { rows: readonly [string, string][] }) {
  return (
    <dl className="ds-proof">
      {rows.map(([label, value]) => (
        <div key={label}><dt>{label}</dt><dd>{value}</dd></div>
      ))}
    </dl>
  )
}

const GRID_COLS = ['A', 'B', 'C', 'D']
const GRID_ROWS = [1, 2, 3, 4, 5]
const GRID_CELLS: Record<string, string> = {
  A1: 'Item', B1: 'Owner', C1: 'Status', D1: 'Score',
  A2: 'Native XLSX', B2: 'Mira', C2: 'Live', D2: '42',
  A3: 'Charts', B3: 'Noah', C3: 'Ready', D3: '18',
  A4: 'PDF', B4: 'Rin', C4: 'Watch', D4: '9',
}

function MiniGrid({
  active,
  onSelect,
  ghost,
}: {
  active: string
  onSelect: (ref: string) => void
  ghost?: string
}) {
  const col = active[0]
  const row = Number(active.slice(1))
  return (
    <div className="ds-grid" role="grid">
      <div className="ds-grid-wrap" style={{ gridTemplateColumns: '40px repeat(4, minmax(72px, 1fr))' }}>
        <div className="ds-corner" />
        {GRID_COLS.map((letter) => (
          <div className="ds-colh" data-active={letter === col ? 'true' : undefined} key={letter}>{letter}</div>
        ))}
        {GRID_ROWS.map((n) => (
          <div key={n} style={{ display: 'contents' }}>
            <div className="ds-rowh" data-active={n === row ? 'true' : undefined}>{n}</div>
            {GRID_COLS.map((letter) => {
              const ref = `${letter}${n}`
              return (
                <button
                  type="button"
                  role="gridcell"
                  key={ref}
                  className="ds-cell"
                  data-selected={ref === active ? 'true' : undefined}
                  data-in-range={ref === ghost ? 'true' : undefined}
                  onClick={() => onSelect(ref)}
                >
                  {GRID_CELLS[ref] ?? ''}
                </button>
              )
            })}
          </div>
        ))}
      </div>
    </div>
  )
}

function MockOverview() {
  return (
    <Frame>
      <header className="ds-appbar">
        <a className="ds-product" href="#/design-system">
          <DsMark />
          <span className="ds-doc-title">
            <strong>InjOffice</strong>
            <small>Source-authoritative file engines</small>
          </span>
        </a>
        <div />
        <div className="ds-appbar-end">
          <DsChip tone="green">Browser engines ready</DsChip>
        </div>
      </header>
      <div className="ds-hero">
        <div>
          <DsChip>Original bytes stay authoritative</DsChip>
          <h3>Edit Office files without rebuilding what you don’t understand.</h3>
          <p>Inspect a real file, apply a bounded change, reopen the exact output, and preserve everything outside the edit.</p>
          <div className="ds-row">
            <DsButton variant="filled">Run the native XLSX proof</DsButton>
            <DsButton variant="outlined">Read the guides</DsButton>
          </div>
          <div className="ds-facts">
            <div><strong>24</strong><span>packages</span></div>
            <div><strong>15</strong><span>proof surfaces</span></div>
            <div><strong>4</strong><span>formats</span></div>
          </div>
        </div>
        <div className="ds-ledger" aria-label="Native XLSX proof sequence">
          <header>
            <div><strong>launch-readiness-plan.xlsx</strong></div>
            <DsChip tone="blue">Browser-local</DsChip>
          </header>
          <ol>
            <li><strong>Extract</strong><code>sha256:bdf753af2b…612f0d</code></li>
            <li><strong>Guard</strong><code>cell.set_value · Data!A1</code></li>
            <li><strong>Apply</strong><code>Go engine in a browser Worker</code></li>
            <li><strong>Verify</strong><code>value confirmed · revision advanced</code></li>
          </ol>
          <footer>
            <DsChip tone="green">Fail closed</DsChip>
            <span className="ds-muted">Unsafe or stale changes produce no replacement file.</span>
          </footer>
        </div>
      </div>
      <div className="ds-cap-list">
        <div className="ds-cap-group">Create and edit</div>
        {[
          ['S', 'Spreadsheets', 'Live workbook or native XLSX mutation'],
          ['D', 'Documents', 'Guarded DOCX text run in browser bytes'],
          ['P', 'Presentations', 'Outline to an editable deck'],
          ['A', 'PDF', 'Page ops, markup, forms, stamps'],
        ].map(([glyph, title, note]) => (
          <div className="ds-cap-row" key={title}>
            <span className="ds-chip">{glyph}</span>
            <span><strong>{title}</strong><small>{note}</small></span>
            <em className="ds-muted">Browser</em>
          </div>
        ))}
      </div>
    </Frame>
  )
}

function MockSheets() {
  const [view, setView] = useState('editor')
  const [active, setActive] = useState('C2')
  const [tab, setTab] = useState('Plan')
  const [collapsed, setCollapsed] = useState(true)

  return (
    <Frame>
      <Head pkg="@injoffice/sheets" title="Spreadsheets" runtime="Browser" />
      <div className="ds-workstrip">
        <DsSegment
          label="Spreadsheet demonstration"
          value={view}
          onChange={setView}
          options={[
            { id: 'editor', label: 'Edit workbook' },
            { id: 'native', label: 'Test XLSX round trip' },
            { id: 'tools', label: 'Package tools' },
          ]}
        />
        <span className="ds-muted">
          {view === 'editor' ? 'Runs in your browser' : view === 'native' ? 'Original bytes stay in this browser' : 'Sparklines, print, outlines, and exchange'}
        </span>
      </div>
      {view === 'editor' && (
        <>
          <div className="ds-formula">
            <input className="ds-namebox" value={active} readOnly aria-label="Active cell" />
            <span className="ds-rule" />
            <span className="ds-fx" aria-hidden="true">fx</span>
            <input value={GRID_CELLS[active] ?? ''} readOnly aria-label="Formula" />
          </div>
          <MiniGrid active={active} onSelect={setActive} />
          <div className="ds-tabbar">
            <button type="button" className="ds-tab-add" aria-label="Add sheet">+</button>
            {['Plan', 'Data', 'Charts'].map((name) => (
              <button key={name} type="button" className="ds-tab" aria-selected={tab === name} onClick={() => setTab(name)}>{name}</button>
            ))}
          </div>
        </>
      )}
      {view === 'native' && (
        <>
          <div className="ds-workstrip">
            <DsButton variant="filled">Use bundled .xlsx</DsButton>
            <DsButton variant="outlined">Open .xlsx</DsButton>
            <DsChip>Browser-local · no upload</DsChip>
          </div>
          <p className="ds-status" data-state="ready">Extracted launch-readiness-plan.xlsx · 3 sheets · revision advanced</p>
          <div className="ds-split">
            <div className="ds-split-main" style={{ padding: 0 }}>
              <MiniGrid active={active} onSelect={setActive} />
            </div>
            <aside className="ds-split-side">
              <div className="ds-panel">
                <span className="ds-eyebrow">01 · Extract</span>
                <Proof rows={[['Revision', 'bdf753af2b…612f0d'], ['Runtime', 'browser-local'], ['Warnings', '0']]} />
              </div>
              <div className="ds-panel">
                <span className="ds-eyebrow">02 · Mutate</span>
                <DsField label="Safe cell">
                  <DsSelect value={active} onChange={(event) => setActive(event.target.value)}>
                    <option value="C2">Plan!C2</option>
                    <option value="D2">Plan!D2</option>
                    <option value="A2">Plan!A2</option>
                  </DsSelect>
                </DsField>
                <DsField label="New literal value">
                  <DsInput defaultValue="Live" />
                </DsField>
                <DsButton variant="green">Save to XLSX</DsButton>
              </div>
              <div className="ds-panel">
                <span className="ds-eyebrow">03 · Verify</span>
                <Proof rows={[['Cell', 'Plan!C2'], ['Before', 'In progress'], ['After', 'Live']]} />
                <DsChip tone="green">Applied</DsChip>
              </div>
            </aside>
          </div>
        </>
      )}
      {view === 'tools' && (
        <div className="ds-cards">
          <div className="ds-panel">
            <span className="ds-eyebrow">@injoffice/sparklines</span>
            <strong>Sparkline</strong>
            <svg className="ds-spark" viewBox="0 0 180 40" aria-hidden="true">
              <polyline fill="none" stroke="var(--ds-select)" strokeWidth="2" points="4,28 34,22 64,24 94,14 124,10 156,6" />
            </svg>
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">@injoffice/print</span>
            <strong>Print workspace</strong>
            <p className="ds-muted">Letter · portrait · 1 page</p>
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">@injoffice/outlines</span>
            <strong>Row groups</strong>
            <div className="ds-tools-grid">
              {['Header', 'North', 'Q1', 'Q2', 'Q3', 'South'].map((label, index) => {
                const hidden = collapsed && index >= 2 && index <= 4
                return (
                  <div key={label} style={{ display: 'contents' }}>
                    <span>{index === 1 ? (collapsed ? '▸' : '▾') : ''}</span>
                    {hidden ? <s>{label}</s> : <span>{label}</span>}
                  </div>
                )
              })}
            </div>
            <DsButton onClick={() => setCollapsed((value) => !value)}>{collapsed ? 'Expand Q1 rows' : 'Collapse Q1 rows'}</DsButton>
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">@injoffice/xlsx-exchange</span>
            <strong>Exchange job</strong>
            <p className="ds-muted">job-1 imported demo.xlsx (4 bytes).</p>
          </div>
        </div>
      )}
    </Frame>
  )
}

function MockDocs() {
  const [draft, setDraft] = useState('Keep Status as Live until native extract finishes.')
  return (
    <Frame>
      <Head pkg="@injoffice/docs" title="Documents" runtime="Browser" />
      <div className="ds-workstrip">
        <DsSegment
          label="DOCX processing runtime"
          value="browser"
          onChange={() => undefined}
          options={[
            { id: 'browser', label: 'In browser (default)' },
            { id: 'server', label: 'Server fallback' },
          ]}
        />
        <DsButton variant="filled">Use bundled .docx</DsButton>
        <DsButton variant="outlined">Open .docx</DsButton>
        <DsChip>Browser-local · no upload</DsChip>
      </div>
      <p className="ds-status" data-state="ready">Extracted northstar-launch-brief.docx · 1 safe text run</p>
      <div className="ds-split">
        <div className="ds-split-main">
          <article className="ds-page">
            <h4>Northstar launch brief</h4>
            <p>Original Office bytes stay authoritative. This preview is a semantic contract, not paginated Word paint.</p>
            <p>{draft}</p>
            <p>Everything outside the selected run stays byte-identical after save.</p>
          </article>
        </div>
        <aside className="ds-split-side">
          <div className="ds-panel">
            <span className="ds-eyebrow">01 · Extract</span>
            <Proof rows={[['Revision', 'c91e2a44…8f10b2'], ['Safe targets', '1'], ['Runtime', 'browser-local']]} />
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">02 · Mutate</span>
            <DsField label="Safe text run">
              <DsSelect defaultValue="r12"><option value="r12">Body · run r12</option></DsSelect>
            </DsField>
            <DsField label="Replacement text">
              <DsTextarea rows={4} value={draft} onChange={(event) => setDraft(event.target.value)} />
            </DsField>
            <DsButton variant="green">Save to DOCX</DsButton>
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">03 · Verify</span>
            <Proof rows={[['Preserved parts', '18 verified'], ['CAS moved', 'c91e… → d04b…']]} />
            <DsChip tone="green">Applied</DsChip>
          </div>
        </aside>
      </div>
    </Frame>
  )
}

function MockSlides() {
  const [at, setAt] = useState(0)
  const slides = [
    { title: 'Northstar launch review', bullets: ['Q3 operating brief', 'Original bytes stay authoritative'] },
    { title: 'Evidence', bullets: ['Native extract in a Worker', 'One guarded mutation', 'Exact-byte readback'] },
    { title: 'Decision', bullets: ['Ship the playground restyle', 'Keep Univer optional'] },
  ]
  const current = slides[at]!
  return (
    <Frame>
      <Head pkg="@injoffice/slides" title="Presentations" runtime="Browser" />
      <div className="ds-workstrip">
        <DsField label="Theme">
          <DsSelect defaultValue="forest">
            <option>forest</option>
            <option>boardroom</option>
            <option>slate</option>
          </DsSelect>
        </DsField>
        <DsField label="Transition">
          <DsSelect defaultValue="fade">
            <option>none</option>
            <option>fade</option>
            <option>push</option>
          </DsSelect>
        </DsField>
        <DsButton variant="filled">Build slides</DsButton>
        <span className="ds-muted">3 slides · built from DeckSpec; no PPTX file is loaded</span>
      </div>
      <div className="ds-split">
        <div className="ds-split-main">
          <div className="ds-slide">
            <h4>{current.title}</h4>
            <ul>{current.bullets.map((item) => <li key={item}>{item}</li>)}</ul>
          </div>
          <textarea className="ds-outline" readOnly value={'Northstar launch review\n- Q3 operating brief\nEvidence\n- Native extract in a Worker'} aria-label="Deck outline" />
        </div>
        <aside className="ds-split-side">
          {slides.map((slide, index) => (
            <button type="button" key={slide.title} className="ds-pick" aria-pressed={at === index} onClick={() => setAt(index)}>
              <strong>{index + 1}. {slide.title}</strong>
            </button>
          ))}
          <div className="ds-panel">
            <span className="ds-eyebrow">Layout QC</span>
            <p className="ds-muted">No estimated overflow or overlap.</p>
          </div>
        </aside>
      </div>
    </Frame>
  )
}

function MockPdf() {
  const [inspector, setInspector] = useState('inspect')
  return (
    <Frame>
      <Head pkg="@injoffice/pdf" title="PDF workbench" runtime="Browser" />
      <div className="ds-workstrip">
        <DsButton variant="filled">Open PDF</DsButton>
        <DsButton>Previous</DsButton>
        <span className="ds-muted">Page 1 of 3</span>
        <DsButton>Next</DsButton>
        <DsField label="Zoom">
          <DsSelect defaultValue="100"><option>100%</option><option>125%</option></DsSelect>
        </DsField>
        <DsButton variant="outlined">Download PDF</DsButton>
      </div>
      <p className="ds-status" data-state="ready">Opened injoffice-sample.pdf · 3 pages</p>
      <div className="ds-split">
        <div className="ds-split-main">
          <article className="ds-pdf-sheet">
            <h4>InjOffice PDF sample</h4>
            <p className="ds-muted">Page 1 · 612×792 pt</p>
            <div className="ds-pdf-line" />
            <div className="ds-pdf-line" />
            <div className="ds-pdf-line ds-pdf-line--short" />
            <div className="ds-pdf-line" />
            <div className="ds-pdf-line ds-pdf-line--short" />
          </article>
        </div>
        <aside className="ds-split-side">
          <div role="tablist" aria-label="PDF tool panels">
            {['inspect', 'pages', 'mark', 'draw', 'forms', 'stamp', 'host'].map((id) => (
              <button
                type="button"
                role="tab"
                key={id}
                className="ds-pick"
                aria-selected={inspector === id}
                aria-pressed={inspector === id}
                onClick={() => setInspector(id)}
              >
                {id === 'host' ? 'Node host' : id[0]!.toUpperCase() + id.slice(1)}
              </button>
            ))}
          </div>
          {inspector === 'inspect' && (
            <div className="ds-panel">
              <span className="ds-eyebrow">Search document</span>
              <DsField label="Query"><DsInput defaultValue="InjOffice" /></DsField>
              <p className="ds-muted">Geometry · 3 pages · 612×792 pt · 0°</p>
            </div>
          )}
          {inspector === 'pages' && (
            <div className="ds-panel">
              <span className="ds-eyebrow">Page operations</span>
              <div className="ds-stack">
                <DsButton variant="outlined">Rotate 90°</DsButton>
                <DsButton variant="outlined">Delete page</DsButton>
                <DsButton variant="outlined">Insert blank after</DsButton>
              </div>
            </div>
          )}
          {inspector === 'mark' && <div className="ds-panel"><span className="ds-eyebrow">Text markup</span><div className="ds-row"><DsButton variant="outlined">highlight</DsButton><DsButton variant="outlined">underline</DsButton></div></div>}
          {inspector === 'draw' && <div className="ds-panel"><span className="ds-eyebrow">Drawings and notes</span><div className="ds-row"><DsButton variant="outlined">Rectangle</DsButton><DsButton variant="outlined">Arrow</DsButton></div></div>}
          {inspector === 'forms' && <div className="ds-panel"><span className="ds-eyebrow">AcroForm values</span><p className="ds-muted">No form fields in this file.</p></div>}
          {inspector === 'stamp' && <div className="ds-panel"><span className="ds-eyebrow">Stamps</span><DsButton variant="outlined">Image stamp</DsButton></div>}
          {inspector === 'host' && <DsCallout tone="note" title="Node host idle">Available during npm run dev. Static Pages builds stay browser-only.</DsCallout>}
        </aside>
      </div>
    </Frame>
  )
}

function MockCharts() {
  const [type, setType] = useState('Column')
  const revenue = [128, 156, 149, 188, 214, 246]
  const max = Math.max(...revenue)
  return (
    <Frame>
      <Head pkg="@injoffice/charts" title="Charts" runtime="Browser" />
      <div className="ds-workstrip">
        <DsField label="Chart type">
          <DsSelect value={type} onChange={(event) => setType(event.target.value)}>
            {['Column', 'Line', 'Area', 'Combination'].map((item) => <option key={item}>{item}</option>)}
          </DsSelect>
        </DsField>
        <label className="ds-check"><input type="checkbox" defaultChecked /> Linear trendline</label>
        <DsButton variant="outlined">Reset data</DsButton>
        <DsButton variant="outlined">Export SVG</DsButton>
        <span className="ds-muted">30 chart types · ChartSpec to ECharts option and native wire</span>
      </div>
      <div className="ds-split ds-split--wide">
        <section className="ds-split-main">
          <span className="ds-eyebrow">Interactive proof</span>
          <div className="ds-row" style={{ marginBottom: 8 }}><h3 style={{ margin: 0, fontSize: 16 }}>Revenue vs target</h3><DsChip tone="blue">{type}</DsChip></div>
          <div className="ds-bars" aria-label={`${type} chart preview`}>
            {revenue.map((value, index) => (
              <div key={index} className="ds-bar" style={{ height: `${Math.round((value / max) * 100)}%` }} />
            ))}
          </div>
          <div className="ds-chart-legend"><span><i />Revenue</span><span><i className="target" />Target</span></div>
        </section>
        <section className="ds-split-side">
          <span className="ds-eyebrow">Source range</span>
          <table className="ds-table">
            <thead><tr><th>Month</th><th className="num">Revenue</th></tr></thead>
            <tbody>
              {['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'].map((month, index) => (
                <tr key={month}><td>{month}</td><td className="num">{revenue[index]}</td></tr>
              ))}
            </tbody>
          </table>
        </section>
      </div>
    </Frame>
  )
}

function MockPivots() {
  return (
    <Frame>
      <Head pkg="@injoffice/pivots" title="Pivot tables" runtime="Browser" />
      <div className="ds-workstrip">
        <DsField label="Group rows"><DsSelect defaultValue="Region"><option>Region</option><option>Owner</option></DsSelect></DsField>
        <DsField label="Value"><DsSelect defaultValue="Revenue"><option>Revenue</option><option>Units</option></DsSelect></DsField>
        <DsField label="Aggregation"><DsSelect defaultValue="sum"><option>sum</option><option>count</option></DsSelect></DsField>
        <label className="ds-check"><input type="checkbox" defaultChecked /> Split by quarter</label>
      </div>
      <div className="ds-split ds-split--wide">
        <section className="ds-split-main">
          <span className="ds-eyebrow">Input · 6 rows</span>
          <table className="ds-table">
            <thead><tr><th>Region</th><th>Owner</th><th>Quarter</th><th className="num">Revenue</th></tr></thead>
            <tbody>
              <tr><td>West</td><td>Maya</td><td>Q1</td><td className="num">42,000</td></tr>
              <tr><td>East</td><td>Theo</td><td>Q1</td><td className="num">37,000</td></tr>
              <tr><td>West</td><td>Maya</td><td>Q2</td><td className="num">51,000</td></tr>
            </tbody>
          </table>
        </section>
        <section className="ds-split-side">
          <span className="ds-eyebrow">Live result · 4 × 4</span>
          <table className="ds-table">
            <thead><tr><th>Region</th><th className="num">Q1</th><th className="num">Q2</th><th className="num">Total</th></tr></thead>
            <tbody>
              <tr><td>East</td><td className="num">37,000</td><td className="num">46,000</td><td className="num">83,000</td></tr>
              <tr><td>West</td><td className="num">42,000</td><td className="num">51,000</td><td className="num">93,000</td></tr>
            </tbody>
          </table>
        </section>
      </div>
    </Frame>
  )
}

const SHAPES = ['roundRect', 'ellipse', 'triangle', 'diamond', 'rightArrow', 'star5', 'hexagon', 'rect'] as const

function ShapeIcon({ kind, size = 32 }: { kind: (typeof SHAPES)[number]; size?: number }) {
  const fill = 'var(--ds-select-soft)'
  const stroke = 'var(--ds-select)'
  return (
    <svg width={size} height={Math.round(size * 0.75)} viewBox="0 0 40 40" fill={fill} stroke={stroke} strokeWidth="2" aria-hidden="true">
      {kind === 'ellipse' && <ellipse cx="20" cy="20" rx="16" ry="12" />}
      {kind === 'triangle' && <polygon points="20,6 36,34 4,34" />}
      {kind === 'diamond' && <polygon points="20,6 36,20 20,34 4,20" />}
      {kind === 'rightArrow' && <polygon points="4,14 22,14 22,8 36,20 22,32 22,26 4,26" />}
      {kind === 'star5' && <polygon points="20,6 24,16 36,16 26,23 30,34 20,27 10,34 14,23 4,16 16,16" />}
      {kind === 'hexagon' && <polygon points="12,8 28,8 36,20 28,32 12,32 4,20" />}
      {kind === 'rect' && <rect x="6" y="10" width="28" height="20" />}
      {kind === 'roundRect' && <rect x="6" y="10" width="28" height="20" rx="6" />}
    </svg>
  )
}

function MockShapes() {
  const [selected, setSelected] = useState<(typeof SHAPES)[number]>('roundRect')
  return (
    <Frame>
      <Head pkg="@injoffice/shapes" title="Shapes" runtime="Browser" />
      <div className="ds-workstrip">
        <DsField label="Category"><DsSelect defaultValue="Basic"><option>Basic</option><option>Arrows</option></DsSelect></DsField>
        <DsField label="Find a shape"><DsInput placeholder="Decision, star, callout…" /></DsField>
        <span className="ds-muted">121 OOXML-native shape kinds</span>
      </div>
      <div className="ds-split">
        <div className="ds-split-main">
          <div className="ds-shape-lib">
            {SHAPES.map((kind) => (
              <button type="button" key={kind} aria-pressed={selected === kind} onClick={() => setSelected(kind)}>
                <ShapeIcon kind={kind} />
                {kind}
              </button>
            ))}
          </div>
        </div>
        <aside className="ds-split-side">
          <span className="ds-eyebrow">Browser preview fidelity</span>
          <div className="ds-shape-stage">
            <ShapeIcon kind={selected} size={96} />
          </div>
          <DsChip tone="green">Distinct browser preview</DsChip>
          <Proof rows={[['OOXML preset', selected], ['Native wire', `Sheet1 ${selected}`]]} />
        </aside>
      </div>
    </Frame>
  )
}

function MockConnectors() {
  const [format, setFormat] = useState('json')
  return (
    <Frame>
      <Head pkg="@injoffice/connectors" title="Data connectors" runtime="Browser" />
      <div className="ds-workstrip">
        <DsSegment
          label="Source format"
          value={format}
          onChange={setFormat}
          options={[{ id: 'json', label: 'JSON' }, { id: 'csv', label: 'CSV' }]}
        />
        {format === 'json' && <DsField label="Dot path"><DsInput defaultValue="data.accounts" /></DsField>}
        <DsButton variant="outlined">Reset sample</DsButton>
        <span className="ds-muted">Host-neutral · credentials never enter the package</span>
      </div>
      <div className="ds-split ds-split--wide">
        <section className="ds-split-main">
          <span className="ds-eyebrow">Host-supplied payload</span>
          <textarea
            className="ds-outline"
            readOnly
            style={{ minHeight: 160 }}
            value={format === 'json'
              ? '{\n  "data": {\n    "accounts": [\n      { "account": "Northwind", "arr": 128000 }\n    ]\n  }\n}'
              : 'account,owner,health,arr\nNorthwind,Maya,Healthy,128000'}
            aria-label={`${format.toUpperCase()} source`}
          />
        </section>
        <section className="ds-split-side">
          <span className="ds-eyebrow">Normalized bound range · 3 rows</span>
          <table className="ds-table">
            <thead><tr><th>account</th><th>owner</th><th className="num">arr</th></tr></thead>
            <tbody>
              <tr><td>Northwind</td><td>Maya</td><td className="num">128,000</td></tr>
              <tr><td>Contoso</td><td>Theo</td><td className="num">94,000</td></tr>
              <tr><td>Globex</td><td>Inez</td><td className="num">151,000</td></tr>
            </tbody>
          </table>
        </section>
      </div>
    </Frame>
  )
}

const FORMULAS = [
  { name: 'XLOOKUP', family: 'Lookup & arrays', formula: '=XLOOKUP(E2, A2:A9, C2:C9)' },
  { name: 'SUMIFS', family: 'Math', formula: '=SUMIFS(D2:D9, B2:B9, "West")' },
  { name: 'FILTER', family: 'Lookup & arrays', formula: '=FILTER(A2:D9, C2:C9="Live")' },
  { name: 'TEXT', family: 'Text', formula: '=TEXT(D2, "$#,##0")' },
  { name: 'IF', family: 'Logic', formula: '=IF(C2="Live", 1, 0)' },
]

function MockFormulas() {
  const [query, setQuery] = useState('')
  const [selected, setSelected] = useState(FORMULAS[0]!)
  const matches = useMemo(
    () => FORMULAS.filter((item) => item.name.toLowerCase().includes(query.trim().toLowerCase())),
    [query],
  )
  return (
    <Frame>
      <Head pkg="@injoffice/formulas" title="Formula coverage" runtime="Browser" />
      <div className="ds-workstrip">
        <DsField label="Search audited functions"><DsInput value={query} onChange={(event) => setQuery(event.target.value)} placeholder="XLOOKUP, date, text…" /></DsField>
        <span className="ds-muted">{matches.length} matching · real-engine audit</span>
      </div>
      <div className="ds-kpis">
        <div><strong>535</strong><span>functions declared</span></div>
        <div><strong>48/48</strong><span>high-value formulas in CI</span></div>
        <div><strong>0</strong><span>#NAME? regressions</span></div>
      </div>
      <div className="ds-split">
        <div className="ds-split-main" style={{ padding: 0 }}>
          {matches.map((item) => (
            <button type="button" key={item.name} className="ds-pick" aria-pressed={selected.name === item.name} onClick={() => setSelected(item)}>
              <strong>{item.name}</strong><span>{item.family}</span>
            </button>
          ))}
        </div>
        <aside className="ds-split-side">
          <span className="ds-eyebrow">Valid fixture invocation</span>
          <h3 style={{ margin: '0 0 8px', fontSize: 16 }}>{selected.name}</h3>
          <p className="ds-code">{selected.formula}</p>
          <DsCallout tone="green" title="Computes in the pinned Univer engine">
            CI evaluates this invocation against a real workbook.
          </DsCallout>
          <DsField label="Calculation job"><DsInput defaultValue="=2+3" /></DsField>
          <DsButton variant="filled">Submit job</DsButton>
        </aside>
      </div>
    </Frame>
  )
}

function MockCollab() {
  const [format, setFormat] = useState('sheets')
  const [left, setLeft] = useState('C2')
  const [right, setRight] = useState('C4')
  return (
    <Frame>
      <Head pkg="@injoffice/collab" title="Collaboration" runtime="Browser" />
      <div className="ds-workstrip">
        <DsSegment
          label="Collaboration demo mode"
          value="simulation"
          onChange={() => undefined}
          options={[
            { id: 'simulation', label: 'Two-editor simulation' },
            { id: 'server', label: 'HTTP + SSE integration' },
          ]}
        />
        <DsSegment
          label="Format"
          value={format}
          onChange={setFormat}
          options={[
            { id: 'sheets', label: 'Sheets' },
            { id: 'slides', label: 'Slides' },
            { id: 'docs', label: 'Docs' },
            { id: 'pdf', label: 'PDF' },
          ]}
        />
        <span className="ds-muted">No server or upload required</span>
      </div>
      <div className="ds-dual">
        <section className="ds-editor-card">
          <header>
            <div className="ds-row"><DsAvatar initials="MN" /><strong>Mira</strong></div>
            <DsChip tone="green">Connected</DsChip>
          </header>
          {format === 'sheets' ? <MiniGrid active={left} onSelect={setLeft} ghost={right} /> : <p className="ds-muted" style={{ padding: 12 }}>Shared {format} canvas for Mira.</p>}
        </section>
        <section className="ds-editor-card">
          <header>
            <div className="ds-row"><DsAvatar initials="NK" tone={2} /><strong>Noah</strong></div>
            <DsChip tone="green">Connected</DsChip>
          </header>
          {format === 'sheets' ? <MiniGrid active={right} onSelect={setRight} ghost={left} /> : <p className="ds-muted" style={{ padding: 12 }}>Shared {format} canvas for Noah.</p>}
        </section>
      </div>
    </Frame>
  )
}

function MockHistory() {
  const [mode, setMode] = useState('grid')
  const [version, setVersion] = useState('v3')
  return (
    <Frame>
      <Head pkg="@injoffice/history" title="History and diffs" runtime="Browser" />
      <div className="ds-workstrip">
        <DsSegment
          label="Diff mode"
          value={mode}
          onChange={setMode}
          options={[{ id: 'grid', label: 'Spreadsheet' }, { id: 'text', label: 'Document' }]}
        />
        <DsButton variant="outlined">Reset example</DsButton>
        <DsButton variant="outlined">Capture after</DsButton>
        <DsButton variant="outlined">Restore selected</DsButton>
      </div>
      <div className="ds-split ds-split--three">
        <section className="ds-split-main">
          <span className="ds-eyebrow">Before</span>
          <textarea className="ds-outline" readOnly value={mode === 'grid' ? 'C2\tIn progress' : 'Status is in progress until extract finishes.'} />
        </section>
        <section className="ds-split-main">
          <span className="ds-eyebrow">After</span>
          <textarea className="ds-outline" readOnly value={mode === 'grid' ? 'C2\tLive' : 'Keep Status as Live until native extract finishes.'} />
        </section>
        <section className="ds-split-main">
          <span className="ds-eyebrow">Changes</span>
          {mode === 'grid' ? (
            <div className="ds-diff-row"><strong>C2</strong><del>In progress</del><span>→</span><ins>Live</ins></div>
          ) : (
            <div className="ds-diff-row"><strong>p1</strong><del>in progress</del><span>→</span><ins>Live</ins></div>
          )}
        </section>
      </div>
      <div className="ds-split ds-split--wide">
        <section className="ds-split-main">
          <span className="ds-eyebrow">Version timeline</span>
          <div className="ds-timeline">
            {[
              ['v1', 'Mira · initial extract'],
              ['v2', 'Agent · cell.set_value'],
              ['v3', 'Mira · user save'],
            ].map(([id, note]) => (
              <button type="button" key={id} aria-pressed={version === id} onClick={() => setVersion(id)}>
                <i />
                <span><strong>{id}</strong><small>{note}</small></span>
              </button>
            ))}
          </div>
        </section>
        <section className="ds-split-side">
          <span className="ds-eyebrow">Isolated preview</span>
          <textarea className="ds-outline" readOnly value={`Preview of ${version}. Restore writes a new immutable version.`} />
        </section>
      </div>
    </Frame>
  )
}

function MockTypography() {
  const [left, setLeft] = useState('Quarterly-')
  const [right, setRight] = useState('results')
  const [size, setSize] = useState(18)
  const decision = left.endsWith('-') ? 'allowed' : 'prohibited'
  return (
    <Frame>
      <Head pkg="@injoffice/font-metrics" title="Typography and layout" runtime="Contract in browser · shaping in Node" />
      <div className="ds-split ds-split--wide">
        <section className="ds-split-main">
          <DsField label="Left cluster"><DsInput value={left} onChange={(event) => setLeft(event.target.value)} /></DsField>
          <DsField label="Right cluster"><DsInput value={right} onChange={(event) => setRight(event.target.value)} /></DsField>
          <DsField label="Font size">
            <input type="range" min={8} max={72} value={size} onChange={(event) => setSize(Number(event.target.value))} />
          </DsField>
          <p className="ds-type-sample" style={{ fontSize: size }}>
            <span>{left}</span><span className="ds-boundary">¦</span><span>{right}</span>
          </p>
        </section>
        <aside className="ds-split-side">
          <DsChip tone={decision === 'allowed' ? 'green' : 'refuse'}>{decision}</DsChip>
          <p className="ds-muted">
            {decision === 'allowed'
              ? 'A native layout engine may wrap at this cluster boundary.'
              : 'The clusters must stay together at this boundary.'}
          </p>
          <Proof rows={[['Run contract', 'valid'], ['Input units', `${size * 1000} milli-pt`], ['Direction', 'LTR / Latn / en-US']]} />
          <DsCallout tone="note" title="Host runtime boundary">
            System-font discovery and HarfBuzz shaping require Node. No browser font measurement is used here.
          </DsCallout>
        </aside>
      </div>
    </Frame>
  )
}

function MockPptxAuthored() {
  const [theme, setTheme] = useState('forest')
  return (
    <Frame>
      <Head pkg="@injoffice/pptx-authored" title="PPTX authoring" runtime="Runs in this browser" />
      <div className="ds-split ds-split--wide">
        <form className="ds-split-main" onSubmit={(event) => event.preventDefault()}>
          <DsField label="Deck title"><DsInput defaultValue="Northstar launch review" /></DsField>
          <DsField label="Subtitle"><DsTextarea rows={3} defaultValue="Q3 operating brief · September 2026" /></DsField>
          <DsField label="Authored theme">
            <DsSelect value={theme} onChange={(event) => setTheme(event.target.value)}>
              {['boardroom', 'midnight', 'slate', 'forest'].map((item) => <option key={item}>{item}</option>)}
            </DsSelect>
          </DsField>
          <p className="ds-muted">Change any field to compile a new immutable native deck. Writing real .pptx bytes remains the Go patcher’s job.</p>
        </form>
        <aside className="ds-split-side">
          <DsChip tone="green">Contract accepted</DsChip>
          <Proof rows={[
            ['Document', 'northstar-launch-review'],
            ['Origin', 'authored'],
            ['Slides', '4'],
            ['Theme', theme],
            ['Compatibility', 'editable'],
          ]} />
          <p className="ds-code">title · body · shape · connector</p>
        </aside>
      </div>
    </Frame>
  )
}

function MockPptxNative() {
  return (
    <Frame>
      <Head pkg="@injoffice/pptx-native" title="Native PPTX model" runtime="Browser" />
      <div className="ds-workstrip">
        <DsSegment
          label="PPTX processing runtime"
          value="browser"
          onChange={() => undefined}
          options={[
            { id: 'browser', label: 'In browser (default)' },
            { id: 'server', label: 'Server fallback' },
          ]}
        />
        <DsButton variant="filled">Use bundled .pptx</DsButton>
        <DsButton variant="outlined">Open .pptx</DsButton>
      </div>
      <p className="ds-status" data-state="ready">Extracted northstar-launch-review.pptx · 3 slides · 4 exact targets</p>
      <div className="ds-split">
        <div className="ds-split-main">
          <table className="ds-table">
            <thead><tr><th>Slide</th><th>Exact target</th><th>Operation</th><th>Current value</th></tr></thead>
            <tbody>
              <tr><td>1</td><td>Title</td><td>text.replace</td><td>Northstar launch review</td></tr>
              <tr><td>2</td><td>Metric diamond</td><td>autoshape.update</td><td>diamond</td></tr>
            </tbody>
          </table>
        </div>
        <aside className="ds-split-side">
          <div className="ds-panel">
            <span className="ds-eyebrow">01 · Extract</span>
            <Proof rows={[['Revision', 'aa19c3…44e1'], ['Warnings', '0']]} />
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">02 · Mutate</span>
            <DsField label="New first-run text"><DsInput defaultValue="Northstar launch review" /></DsField>
            <DsButton variant="green">Save to PPTX</DsButton>
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">03 · Verify</span>
            <DsChip tone="green">Untouched anchors verified</DsChip>
          </div>
        </aside>
      </div>
    </Frame>
  )
}

function MockPptxRender() {
  const [preset, setPreset] = useState('roundRect')
  const commands = ['beginSlide', 'fillPath', 'strokePath', 'drawConnector', 'endSlide']
  return (
    <Frame>
      <Head pkg="@injoffice/pptx-render" title="PPTX renderer" runtime="Browser-safe command layer" />
      <div className="ds-split ds-split--wide">
        <section className="ds-split-main">
          <DsField label="Native shape preset">
            <DsSelect value={preset} onChange={(event) => setPreset(event.target.value)}>
              {['rect', 'roundRect', 'ellipse', 'triangle'].map((item) => <option key={item}>{item}</option>)}
            </DsSelect>
          </DsField>
          <label className="ds-check"><input type="checkbox" defaultChecked /> Include a connector in the command stream</label>
          <div className="ds-shape-stage">
            <svg width="160" height="90" viewBox="0 0 160 90" aria-label={`${preset} native preset path`}>
              {preset === 'ellipse'
                ? <ellipse cx="70" cy="45" rx="48" ry="28" fill="var(--ds-select-soft)" stroke="var(--ds-select)" strokeWidth="3" />
                : <rect x="22" y="18" width="96" height="54" rx={preset === 'roundRect' ? 12 : 0} fill="var(--ds-select-soft)" stroke="var(--ds-select)" strokeWidth="3" />}
              <line x1="118" y1="45" x2="150" y2="70" stroke="var(--ds-refuse)" strokeWidth="3" />
            </svg>
          </div>
        </section>
        <aside className="ds-split-side">
          <DsChip tone="green">Recording complete</DsChip>
          <strong>{commands.length} commands</strong>
          {commands.map((command) => (
            <div className="ds-command" key={command}><code>{command}</code><span className="ds-muted">demo-shape</span></div>
          ))}
        </aside>
      </div>
    </Frame>
  )
}

function MockGuides() {
  const [page, setPage] = useState('quickstart')
  return (
    <Frame>
      <header className="ds-appbar">
        <a className="ds-product" href="#/design-system">
          <DsMark />
          <span className="ds-doc-title">
            <strong>InjOffice</strong>
            <small>Documentation</small>
          </span>
        </a>
        <input className="ds-guides-search" placeholder="Search guides and packages" aria-label="Search documentation" />
        <div className="ds-appbar-end">
          <DsButton>Showcase</DsButton>
          <DsButton>Packages</DsButton>
          <DsButton variant="outlined">Workbench</DsButton>
        </div>
      </header>
      <div className="ds-guides">
        <nav className="ds-guides-nav" aria-label="Guide navigation">
          <span className="ds-eyebrow">Start</span>
          {['introduction', 'installation', 'quickstart'].map((id) => (
            <button type="button" key={id} aria-current={page === id ? 'true' : undefined} onClick={() => setPage(id)}>
              {id[0]!.toUpperCase() + id.slice(1)}
            </button>
          ))}
          <span className="ds-eyebrow">Native files</span>
          {['xlsx', 'docx', 'pptx'].map((id) => (
            <button type="button" key={id} aria-current={page === id ? 'true' : undefined} onClick={() => setPage(id)}>{id.toUpperCase()}</button>
          ))}
        </nav>
        <article className="ds-guides-article">
          <DsChip tone="green">Apache-2.0</DsChip>
          <h3>{page === 'quickstart' ? 'Apply one guarded change' : page.toUpperCase()}</h3>
          <p className="ds-muted">Install the package, copy the snippet, then run the same proof the playground uses.</p>
          <pre>{`import { extractWorkbook } from '@injoffice/xlsx-wasm'\n\nconst workbook = await extractWorkbook(bytes)`}</pre>
          <DsButton variant="filled">Open the working proof</DsButton>
        </article>
      </div>
    </Frame>
  )
}

function MockAgent() {
  const [tool, setTool] = useState<'sheets' | 'docs' | 'slides' | 'pdf'>('sheets')
  const labels = { sheets: 'Sheets', docs: 'Docs', slides: 'Slides', pdf: 'PDF' } as const
  const preview = {
    sheets: {
      change: 'xlsx.cell.set_value · Forecast!D5 · High',
      body: (
        <table className="ds-table">
          <thead><tr><th>Region</th><th>Plan</th><th>Actual</th><th>Confidence</th></tr></thead>
          <tbody>
            <tr><td>North</td><td className="num">120</td><td className="num">126</td><td>High</td></tr>
            <tr><td>East</td><td className="num">80</td><td className="num">84</td><td className="num">High</td></tr>
          </tbody>
        </table>
      ),
    },
    docs: {
      change: 'docx.text.replace · block-summary',
      body: (
        <article className="ds-page">
          <h4>Northstar launch brief</h4>
          <p>Launch readiness is on track for the October 18 review.</p>
          <p>Owner: Product Operations</p>
        </article>
      ),
    },
    slides: {
      change: 'pptx.authored.slide.update · 86% → 91%',
      body: (
        <div className="ds-slide">
          <h4>Launch readiness</h4>
          <ul>
            <li>Five regions prepared</li>
            <li>Readiness 91%</li>
          </ul>
        </div>
      ),
    },
    pdf: {
      change: 'pdf.page.rotate · page 2 · 90°',
      body: (
        <article className="ds-pdf-sheet">
          <h4>Approval record</h4>
          <p className="ds-muted">Review packet · page 2 · rotation 90°</p>
          <div className="ds-pdf-line" />
          <div className="ds-pdf-line ds-pdf-line--short" />
          <div className="ds-pdf-line" />
        </article>
      ),
    },
  }[tool]
  return (
    <Frame>
      <Head pkg="@injoffice/agent-tools" title={`AI change sets · ${labels[tool]}`} runtime="Browser" />
      <div className="ds-workstrip">
        <DsSegment
          label="Office tool"
          value={tool}
          onChange={(id) => setTool(id as 'sheets' | 'docs' | 'slides' | 'pdf')}
          options={[
            { id: 'sheets', label: 'Sheets' },
            { id: 'docs', label: 'Docs' },
            { id: 'slides', label: 'Slides' },
            { id: 'pdf', label: 'PDF' },
          ]}
        />
        <span className="ds-muted">Inspect → plan → preview → validate → host approval → commit → verify</span>
      </div>
      <div className="ds-split">
        <div className="ds-split-main">
          <span className="ds-eyebrow">Isolated preview</span>
          {preview.body}
        </div>
        <aside className="ds-split-side">
          <div className="ds-panel">
            <span className="ds-eyebrow">Change set</span>
            <p className="ds-code">{preview.change}</p>
          </div>
          <div className="ds-panel">
            <span className="ds-eyebrow">Validate</span>
            <DsChip tone="green">Revision bound</DsChip>
            <p className="ds-muted">Host approval is required before commit. A stale revision refuses and writes nothing.</p>
          </div>
          <DsButton variant="filled">Approve and commit</DsButton>
          <DsButton variant="refuse">Refuse</DsButton>
        </aside>
      </div>
    </Frame>
  )
}

export const SURFACE_GALLERY: {
  id: string
  label: string
  blurb: string
  Mock: () => ReactNode
}[] = [
  { id: 'surface-overview', label: 'Overview', blurb: 'Hero, native proof ledger, and the fifteen labs grouped by job.', Mock: MockOverview },
  { id: 'surface-sheets', label: 'Spreadsheets', blurb: 'Editor, native XLSX round trip, and package tools behind one switcher.', Mock: MockSheets },
  { id: 'surface-docs', label: 'Documents', blurb: 'Semantic DOCX preview with extract, mutate, and verify in the inspector.', Mock: MockDocs },
  { id: 'surface-slides', label: 'Presentations', blurb: 'Outline to canvas. Theme, transition, and layout QC — no PPTX file loaded.', Mock: MockSlides },
  { id: 'surface-pdf', label: 'PDF', blurb: 'Page stage plus inspect, pages, markup, draw, forms, stamp, and host tools.', Mock: MockPdf },
  { id: 'surface-charts', label: 'Charts', blurb: 'Edit a source range and preview one of the chart types as a stage, not a widget kit.', Mock: MockCharts },
  { id: 'surface-pivots', label: 'Pivot tables', blurb: 'Source records on the left, baked pivot on the right, wells in the toolbar.', Mock: MockPivots },
  { id: 'surface-shapes', label: 'Shapes', blurb: 'Preset library and a labeled preview. Approximate kinds stay honest.', Mock: MockShapes },
  { id: 'surface-connectors', label: 'Data connectors', blurb: 'JSON or CSV in, bound grid out. Credentials never enter the package.', Mock: MockConnectors },
  { id: 'surface-formulas', label: 'Formulas', blurb: 'Search the audited set, inspect a fixture invocation, submit a calculation job.', Mock: MockFormulas },
  { id: 'surface-agent', label: 'AI change sets', blurb: 'Inspect, plan, preview, validate, and commit a bounded Sheets, Docs, Slides, or PDF change set. The host owns approval.', Mock: MockAgent },
  { id: 'surface-collab', label: 'Collaboration', blurb: 'Two independent editors, presence, and a format switch. Simulation first.', Mock: MockCollab },
  { id: 'surface-history', label: 'History', blurb: 'Before, after, structured diffs, and an immutable version timeline.', Mock: MockHistory },
  { id: 'surface-font-metrics', label: 'Typography', blurb: 'Browser-safe layout contract. Shaping stays on the Node side of the line.', Mock: MockTypography },
  { id: 'surface-pptx-authored', label: 'PPTX authoring', blurb: 'DeckSpec fields compile into a native contract without a DOM.', Mock: MockPptxAuthored },
  { id: 'surface-pptx-native', label: 'Native PPTX', blurb: 'Exact text or AutoShape targets on real bytes, then verify and download.', Mock: MockPptxNative },
  { id: 'surface-pptx-render', label: 'PPTX renderer', blurb: 'Render tree to an ordered paint stream. The SVG is a host preview only.', Mock: MockPptxRender },
  { id: 'surface-guides', label: 'Guides', blurb: 'Docs shell: search, grouped nav, article, and a copy-paste snippet.', Mock: MockGuides },
]
