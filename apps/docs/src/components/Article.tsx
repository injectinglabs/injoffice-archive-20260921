import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { NAV_GROUPS, PAGES, PAGE_BY_ID, type DocId } from '../catalog'

export function Article({
  id,
  children,
}: {
  id: DocId
  children: ReactNode
}) {
  const page = PAGE_BY_ID.get(id)!
  const ref = useRef<HTMLElement>(null)
  const [toc, setToc] = useState<{ id: string; text: string }[]>([])
  const [active, setActive] = useState<string>()

  useEffect(() => {
    const root = ref.current
    if (!root) return
    const headings = [...root.querySelectorAll('h2[id]')].map((node) => ({
      id: node.id,
      text: node.textContent ?? '',
    }))
    setToc(headings)
    const observer = new IntersectionObserver((entries) => {
      const visible = entries.filter((entry) => entry.isIntersecting).at(-1)
      if (visible?.target.id) setActive(visible.target.id)
    }, { rootMargin: '-80px 0px -60% 0px' })
    for (const node of root.querySelectorAll('h2[id]')) observer.observe(node)
    return () => observer.disconnect()
  }, [id, children])

  const { prev, next } = useMemo(() => {
    const sequence = PAGES.filter((item) => NAV_GROUPS.includes(item.group as typeof NAV_GROUPS[number]))
    const index = sequence.findIndex((item) => item.id === id)
    return { prev: sequence[index - 1], next: sequence[index + 1] }
  }, [id])

  return (
    <>
      <article className="article" ref={ref}>
        <div className="kicker">
          {page.packageName ? <span className="badge badge--pkg">{page.packageName}</span> : null}
          {page.runtime ? <span className="badge">{page.runtime}</span> : null}
          {page.authority ? (
            <span className={`badge ${page.authority.includes('OOXML') || page.authority.includes('PDF') ? 'badge--pass' : page.authority.includes('Univer') ? 'badge--patch' : ''}`}>
              {page.authority}
            </span>
          ) : null}
        </div>
        <h1>{page.title}</h1>
        <p className="lead">{page.description}</p>
        {children}
        <nav className="pager" aria-label="Adjacent guides">
          {prev ? <a href={prev.href}><small>Previous</small>{prev.navTitle}</a> : <span />}
          {next ? <a href={next.href} style={{ textAlign: 'right' }}><small>Next</small>{next.navTitle}</a> : <span />}
        </nav>
      </article>
      <aside className="toc" aria-label="On this page">
        <strong>On this page</strong>
        {toc.map((item) => (
          <a key={item.id} href={`${page.href}?section=${item.id}`} aria-current={active === item.id ? 'true' : undefined}>{item.text}</a>
        ))}
      </aside>
    </>
  )
}
