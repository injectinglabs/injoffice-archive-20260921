import { PAGES, type DocId } from './catalog'

export const DOCS_ROOT = '#/guides'

export function isDocsHash(hash = typeof location === 'undefined' ? '' : location.hash): boolean {
  const path = hash.replace(/^#\/?/, '').split(/[/?]/)[0]?.toLowerCase() ?? ''
  return path === 'guides'
}

export function parsePath(hash = typeof location === 'undefined' ? '' : location.hash): DocId {
  const raw = hash.replace(/^#/, '')
  const path = (raw.split('?')[0] || '/').replace(/\/$/, '') || '/'
  if (path === '/guides' || path === 'guides') return 'home'
  const href = `#${path.startsWith('/') ? path : `/${path}`}`
  const match = PAGES.find((page) => page.href === href)
  return match?.id ?? 'home'
}

export function parseSection(hash = typeof location === 'undefined' ? '' : location.hash): string | null {
  const query = hash.split('?')[1] ?? ''
  return new URLSearchParams(query).get('section')
}

export function navigate(href: string) {
  if (href.startsWith('#/guides')) {
    const current = `${location.hash}`
    if (current.split('?')[0] === href.split('?')[0] && href.includes('section=')) {
      location.hash = href
      return
    }
    location.hash = href
    if (!href.includes('section=')) window.scrollTo(0, 0)
    return
  }
  location.hash = href.startsWith('#') ? href : `#/${href.replace(/^\//, '')}`
}
