import { useMemo, useState, type MouseEvent, type ReactNode } from 'react'
import { NAV_GROUPS, PAGES, PAGE_BY_ID, type DocId } from '../catalog'
import { navigate } from '../route'

function handleNav(event: MouseEvent<HTMLAnchorElement>) {
  const href = event.currentTarget.getAttribute('href')
  if (!href || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return
  event.preventDefault()
  navigate(href)
}

export function Layout({
  id,
  children,
}: {
  id: DocId
  children: ReactNode
}) {
  const [query, setQuery] = useState('')
  const [open, setOpen] = useState(false)
  const page = PAGE_BY_ID.get(id)!
  const hits = useMemo(() => {
    const needle = query.trim().toLowerCase()
    if (!needle) return []
    return PAGES.filter((item) => item.id !== 'home' && (
      item.title.toLowerCase().includes(needle)
      || item.description.toLowerCase().includes(needle)
      || item.keywords.some((keyword) => keyword.includes(needle))
      || (item.packageName?.includes(needle) ?? false)
    )).slice(0, 8)
  }, [query])

  return (
    <div className="inj-docs">
      <a className="skip" href="#doc-main">Skip to content</a>
      <header className="topbar">
        <a className="brand" href="#/guides" onClick={handleNav}>
          <img src="/logo.svg" alt="" />
          <div>
            <strong>InjOffice</strong>
            <small>Documentation</small>
          </div>
        </a>
        <label className="top-search">
          <span className="visually-hidden">Search</span>
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Search guides and packages"
            aria-label="Search documentation"
          />
          <kbd>/</kbd>
        </label>
        <nav className="top-links" aria-label="Docs sections">
          <button className="menu-btn" type="button" onClick={() => setOpen((value) => !value)} aria-expanded={open}>Menu</button>
          <a href="#/guides/showcase" onClick={handleNav} aria-current={id === 'showcase' ? 'page' : undefined}>Showcase</a>
          <a href="#/guides/reference" onClick={handleNav} aria-current={id === 'reference' ? 'page' : undefined}>Packages</a>
          <a href="#/overview">Workbench</a>
          <a href="https://github.com/injectinglabs/injoffice" target="_blank" rel="noreferrer">GitHub</a>
        </nav>
      </header>
      {hits.length > 0 && (
        <div className="search-pop" role="listbox">
          {hits.map((item) => (
            <a key={item.id} href={item.href} onClick={(event) => { handleNav(event); setQuery('') }}>
              <strong>{item.navTitle}</strong>
              <span>{item.packageName ?? item.group} · {item.description}</span>
            </a>
          ))}
        </div>
      )}
      <div className="shell">
        <aside className={`sidebar${open ? ' open' : ''}`} aria-label="Guide navigation">
          {NAV_GROUPS.map((group) => (
            <section className="nav-group" key={group}>
              <h2>{group}</h2>
              {PAGES.filter((item) => item.group === group).map((item) => (
                <a key={item.id} href={item.href} aria-current={item.id === id ? 'page' : undefined} onClick={(event) => { handleNav(event); setOpen(false) }}>
                  {item.navTitle}
                </a>
              ))}
            </section>
          ))}
        </aside>
        <main className="main" id="doc-main" data-page={page.id}>
          {children}
        </main>
      </div>
    </div>
  )
}
