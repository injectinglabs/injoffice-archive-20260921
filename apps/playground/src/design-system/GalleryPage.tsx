import { useEffect, useMemo, useState } from 'react'
import { currentColorScheme, persistColorScheme, type ColorScheme } from '../colorScheme'
import './tokens.css'
import './components.css'
import {
  DsAvatar,
  DsButton,
  DsCallout,
  DsChip,
  DsField,
  DsInput,
  DsMark,
  DsSelect,
  DsTextarea,
  DsTool,
  IconBorder,
  IconFilter,
  IconPrint,
  IconRedo,
  IconUndo,
} from './primitives'
import { SURFACE_GALLERY } from './surfaces'

const CHROME_SECTIONS = [
  { id: 'intent', label: 'Intent' },
  { id: 'workbook', label: 'Workbook chrome' },
  { id: 'grid', label: 'Grid' },
  { id: 'color', label: 'Color' },
  { id: 'type', label: 'Type' },
  { id: 'buttons', label: 'Buttons' },
  { id: 'inputs', label: 'Inputs' },
  { id: 'dialogs', label: 'Dialogs' },
  { id: 'status', label: 'Status' },
  { id: 'presence', label: 'Invite and comments' },
]

const COLS = ['A', 'B', 'C', 'D', 'E']
const ROWS = [1, 2, 3, 4, 5, 6]
const CELLS: Record<string, string> = {
  A1: 'Item', B1: 'Owner', C1: 'Status', D1: 'Score',
  A2: 'Native XLSX', B2: 'Mira', C2: 'Live', D2: '42',
  A3: 'Charts', B3: 'Noah', C3: 'Ready', D3: '18',
  A4: 'Count', D4: '=COUNTA(A2:A3)',
}

const MENUS = ['File', 'Edit', 'View', 'Insert', 'Format', 'Data', 'Tools']
const SWATCHES: [string, string, string][] = [
  ['Paper', 'var(--ds-paper)', 'Worksheet surface.'],
  ['Chrome', 'var(--ds-chrome)', 'Toolbar and tab bar.'],
  ['Ink', 'var(--ds-ink)', 'Menus, cells, titles.'],
  ['Grid', 'var(--ds-grid)', 'Cell hairlines.'],
  ['Select', 'var(--ds-select)', 'Active cell and primary actions.'],
  ['Forest', 'var(--ds-green)', 'Selected sheet tab and applied state.'],
  ['Refuse', 'var(--ds-refuse)', 'Fail-closed. Original bytes unchanged.'],
]

function parseSection(hash = location.hash): string | null {
  return new URLSearchParams(hash.split('?')[1] ?? '').get('section')
}

function MiniWorkbook() {
  const [active, setActive] = useState('C2')
  const [tab, setTab] = useState('Plan')
  const [formula, setFormula] = useState(CELLS.C2)
  const [title, setTitle] = useState('launch-readiness-plan')

  const col = active[0]
  const row = Number(active.slice(1))
  const cols = useMemo(() => COLS, [])
  const rows = useMemo(() => ROWS, [])

  const select = (ref: string) => {
    setActive(ref)
    setFormula(CELLS[ref] ?? '')
  }

  return (
    <div className="ds-sheet-frame">
      <header className="ds-appbar">
        <a className="ds-product" href="#/design-system">
          <DsMark />
          <span className="ds-doc-title">
            <input value={title} onChange={(event) => setTitle(event.target.value)} aria-label="Workbook name" />
            <small>InjOffice</small>
          </span>
        </a>
        <div />
        <div className="ds-appbar-end">
          <div className="ds-people" aria-label="Editors">
            <DsAvatar initials="MN" />
            <DsAvatar initials="AK" tone={2} />
          </div>
          <DsButton variant="filled">Invite</DsButton>
        </div>
      </header>
      <nav className="ds-menubar" aria-label="Workbook menus">
        {MENUS.map((menu) => <button type="button" key={menu}>{menu}</button>)}
      </nav>
      <div className="ds-toolbar" role="toolbar" aria-label="Formatting">
        <DsTool aria-label="Undo"><IconUndo /></DsTool>
        <DsTool aria-label="Redo"><IconRedo /></DsTool>
        <DsTool aria-label="Print"><IconPrint /></DsTool>
        <span className="ds-sep" />
        <DsTool aria-label="Bold"><strong>B</strong></DsTool>
        <DsTool aria-label="Italic"><em>I</em></DsTool>
        <span className="ds-sep" />
        <DsTool aria-label="Currency">$</DsTool>
        <DsTool aria-label="Percent">%</DsTool>
        <span className="ds-sep" />
        <DsTool aria-label="Filter"><IconFilter /></DsTool>
        <DsTool aria-label="Borders"><IconBorder /></DsTool>
      </div>
      <div className="ds-formula">
        <input className="ds-namebox" value={active} readOnly aria-label="Active cell" />
        <span className="ds-rule" />
        <span className="ds-fx" aria-hidden="true">fx</span>
        <input value={formula} onChange={(event) => setFormula(event.target.value)} aria-label="Formula" />
      </div>
      <div className="ds-grid" role="grid" aria-label="Sample worksheet">
        <div
          className="ds-grid-wrap"
          style={{ gridTemplateColumns: `${46}px repeat(${cols.length}, minmax(100px, 1fr))` }}
        >
          <div className="ds-corner" />
          {cols.map((letter) => (
            <div className="ds-colh" data-active={letter === col ? 'true' : undefined} key={letter}>{letter}</div>
          ))}
          {rows.map((n) => (
            <div key={n} style={{ display: 'contents' }}>
              <div className="ds-rowh" data-active={n === row ? 'true' : undefined}>{n}</div>
              {cols.map((letter) => {
                const ref = `${letter}${n}`
                return (
                  <button
                    type="button"
                    role="gridcell"
                    key={ref}
                    className="ds-cell"
                    data-selected={ref === active ? 'true' : undefined}
                    data-in-range={letter === 'C' && n === 2 && ref !== active ? 'true' : undefined}
                    onClick={() => select(ref)}
                  >
                    {CELLS[ref] ?? ''}
                  </button>
                )
              })}
            </div>
          ))}
        </div>
      </div>
      <div className="ds-tabbar">
        <button type="button" className="ds-tab-add" aria-label="Add sheet">+</button>
        {['Plan', 'Data', 'Charts'].map((name) => (
          <button
            key={name}
            type="button"
            className="ds-tab"
            aria-selected={tab === name}
            onClick={() => setTab(name)}
          >
            {name}
          </button>
        ))}
      </div>
    </div>
  )
}

export default function GalleryPage() {
  const [scheme, setScheme] = useState<ColorScheme>(() => currentColorScheme())
  const [section, setSection] = useState<string | null>(() => parseSection())
  const [bold, setBold] = useState(true)

  useEffect(() => {
    const sync = () => setSection(parseSection())
    window.addEventListener('hashchange', sync)
    return () => window.removeEventListener('hashchange', sync)
  }, [])

  useEffect(() => {
    if (!section) return
    document.getElementById(section)?.scrollIntoView({ block: 'start' })
  }, [section])

  const setTheme = (next: ColorScheme) => {
    persistColorScheme(next)
    setScheme(next)
  }

  return (
    <div className="ds">
      <div className="ds-gallery">
        <main className="ds-gallery-main">
          <p className="ds-kicker">
            <DsChip tone="green">Applied</DsChip>
            <DsChip>Live playground and guides</DsChip>
          </p>
          <h1>A spreadsheet workbench, a little more current.</h1>
          <p className="ds-lede">
            Same bones: formula bar, column letters, active cell, sheet tabs. Rounder chrome, airier rows, Instrument Sans, pill tabs, and a rounded grid mark — still not a product logo. The live playground and guides now use this language. Toggle light and dark here, then open the demo.
          </p>
          <div className="ds-row" style={{ marginBottom: 20 }}>
            <DsButton variant={scheme === 'light' ? 'outlined' : 'text'} onClick={() => setTheme('light')}>Light</DsButton>
            <DsButton variant={scheme === 'dark' ? 'outlined' : 'text'} onClick={() => setTheme('dark')}>Dark</DsButton>
            <a className="ds-btn ds-btn--outlined" href="#/overview">Current demo</a>
          </div>

          <section id="intent">
            <h2>Intent</h2>
            <p>
              People already know how to read a spreadsheet. The playground should feel like opening a workbook: title in the header, File/Edit menus, formula bar, grid, tabs along the bottom. Applied and refused stay as quiet chips, not a separate visual language.
            </p>
          </section>

          <section id="workbook">
            <h2>Workbook chrome</h2>
            <p>Hairline grid mark, serif title, invite cluster, menus, toolbar, formula bar, grid, and sheet tabs — one stacked sheet.</p>
            <MiniWorkbook />
          </section>

          <section id="grid">
            <h2>Grid</h2>
            <p>
              Sticky row and column headers. Selected header tints blue. The active cell gets a 2px blue ring and a fill handle. Range fill is the pale blue wash. Click cells in the workbook above.
            </p>
          </section>

          <section id="color">
            <h2>Color</h2>
            <p>Paper, hairline grid, forest for the selected tab, steel-blue for the active cell. No extra accent palette.</p>
            <div className="ds-stack">
              {SWATCHES.map(([name, fill, note]) => (
                <div className="ds-swatch" key={name}>
                  <i style={{ background: fill }} />
                  <div>
                    <strong>{name}</strong>
                    <div style={{ color: 'var(--ds-mute)', fontSize: 13 }}>{note}</div>
                  </div>
                </div>
              ))}
            </div>
          </section>

          <section id="type">
            <h2>Type</h2>
            <p>
              Source Serif for the workbook title. Instrument Sans in chrome and cells. IBM Plex Mono for formulas and hashes.
            </p>
            <p style={{ fontFamily: 'var(--ds-display)', fontSize: 28, fontWeight: 500, letterSpacing: '-0.03em', margin: '0 0 8px' }}>launch-readiness-plan</p>
            <p className="ds-lede" style={{ fontFamily: 'var(--ds-cell)' }}>Native XLSX · Mira · Live</p>
            <p className="ds-code">=COUNTA(A2:A3) · sha256:bdf753af2b…612f0d</p>
          </section>

          <section id="buttons">
            <h2>Buttons</h2>
            <p>Text buttons for menus. Filled steel for Invite. Forest only on the selected cell in the mark and the active tab. Outlined for file actions.</p>
            <div className="ds-row">
              <DsButton variant="filled">Invite</DsButton>
              <DsButton variant="green">Apply</DsButton>
              <DsButton>Cancel</DsButton>
              <DsButton variant="outlined">Open .xlsx</DsButton>
              <DsButton variant="refuse">Refuse</DsButton>
              <DsButton disabled>Disabled</DsButton>
            </div>
            <div className="ds-toolbar" style={{ marginTop: 12, border: '1px solid var(--ds-line)' }}>
              <DsTool aria-pressed={bold} onClick={() => setBold((value) => !value)}><strong>B</strong></DsTool>
              <DsTool><em>I</em></DsTool>
              <span className="ds-sep" />
              <DsTool>$</DsTool>
              <DsTool>%</DsTool>
            </div>
          </section>

          <section id="inputs">
            <h2>Inputs</h2>
            <p>Dialog fields are 36px, 4px radius, blue focus ring. The formula bar stays a flat hairline row, not a material text field.</p>
            <div className="ds-field-row">
              <DsField label="Find in sheet">
                <DsInput defaultValue="Native XLSX" />
              </DsField>
              <DsField label="Number format">
                <DsSelect defaultValue="automatic">
                  <option value="automatic">Automatic</option>
                  <option value="number">Number</option>
                  <option value="currency">Currency</option>
                </DsSelect>
              </DsField>
            </div>
            <DsField label="Custom formula">
              <DsTextarea defaultValue="=COUNTA(A2:A3)" rows={3} spellCheck={false} />
            </DsField>
          </section>

          <section id="dialogs">
            <h2>Dialogs</h2>
            <p>White sheet, 6px corners, actions right-aligned. Snackbars are dark bars for transient proof results.</p>
            <div className="ds-dialog">
              <h3>Apply this change?</h3>
              <p>cell.set_value on Data!A1, bound to the current package SHA. Everything outside the edit stays byte-identical.</p>
              <div className="ds-dialog-actions">
                <DsButton>Cancel</DsButton>
                <DsButton variant="filled">Apply</DsButton>
              </div>
            </div>
            <div className="ds-snack" style={{ marginTop: 12 }}>Change applied · revision advanced</div>
          </section>

          <section id="status">
            <h2>Status</h2>
            <p>Chips for runtime and package. Forest wash for landed edits. Red wash for fail-closed.</p>
            <div className="ds-row">
              <DsChip>@injoffice/xlsx-wasm</DsChip>
              <DsChip tone="blue">Browser Worker</DsChip>
              <DsChip tone="green">Applied</DsChip>
              <DsChip tone="refuse">#STALE_REVISION</DsChip>
            </div>
            <div className="ds-stack">
              <DsCallout tone="note" title="Saved in this browser">
                The Worker extracted the bundled workbook. No sidecar.
              </DsCallout>
              <DsCallout tone="green" title="Applied">
                Data!C2 is Live. Package SHA advanced.
              </DsCallout>
              <DsCallout tone="refuse" title="Refused">
                STALE_REVISION. The original bytes were not replaced.
              </DsCallout>
            </div>
          </section>

          <section id="presence">
            <h2>Invite and comments</h2>
            <p>Overlapping initials in the app bar. Notes are a paper card with a select-colored rule, not a sticky.</p>
            <div className="ds-row">
              <div className="ds-people">
                <DsAvatar initials="MN" />
                <DsAvatar initials="AK" tone={2} />
                <DsAvatar initials="JN" tone={3} />
              </div>
              <DsButton variant="filled">Invite</DsButton>
            </div>
            <div className="ds-comment">
              <header>Mira North <span style={{ fontWeight: 400, color: 'var(--ds-mute)' }}>on C2</span></header>
              Keep Status as Live until native extract finishes.
            </div>
          </section>

          <section id="surfaces">
            <h2>Surfaces</h2>
            <p>
              Reference layouts for every playground route. The live demo and guides now share these tokens and controls. `#/docs` remains the DOCX lab.
            </p>
          </section>

          {SURFACE_GALLERY.map((item) => (
            <section id={item.id} key={item.id}>
              <h2>{item.label}</h2>
              <p>{item.blurb}</p>
              <item.Mock />
            </section>
          ))}
        </main>

        <nav className="ds-toc" aria-label="Design system sections">
          <p className="ds-toc-group">Chrome</p>
          {CHROME_SECTIONS.map((item) => (
            <a
              key={item.id}
              href={`#/design-system?section=${item.id}`}
              aria-current={section === item.id ? 'true' : undefined}
            >
              {item.label}
            </a>
          ))}
          <p className="ds-toc-group">Surfaces</p>
          {SURFACE_GALLERY.map((item) => (
            <a
              key={item.id}
              href={`#/design-system?section=${item.id}`}
              aria-current={section === item.id ? 'true' : undefined}
            >
              {item.label}
            </a>
          ))}
        </nav>
      </div>
    </div>
  )
}
