import { readFileSync } from 'node:fs'
import { createElement, Fragment } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { collabRoomHref, initialCollabArtifact } from './collabComposition'
import { parseCollabQuery, type CollabFormat } from './collabScope'
import HistoryPage from './pages/HistoryPage'
import FontMetricsPage from './pages/FontMetricsPage'
import ShapesPage from './pages/ShapesPage'
import { HistoryTimeline } from '../../../packages/history/src/react'
import type { HistoryCommandController } from '../../../packages/history/src/commands'
import { SimEditorFrame, SIM_EDITORS } from './collabSimChrome'
import { CollabRoomChrome } from './collab/roomChrome'

const source = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')
const formats: CollabFormat[] = ['sheets', 'docs', 'slides', 'pdf']

describe('shared panel composition', () => {
  it.each(formats)('does not seed an embedded %s room from another workspace URL', (fixedFormat) => {
    expect(initialCollabArtifact({ fixedFormat }, 'http://localhost:3100/?artifact=unrelated&format=pdf#/pdf')).toBe('')
    expect(initialCollabArtifact({ fixedFormat, initialHash: `#/collab?format=${fixedFormat}&artifact=explicit-room` }, 'http://localhost:3100/?artifact=unrelated')).toBe('explicit-room')
    const differentFormat = fixedFormat === 'docs' ? 'sheets' : 'docs'
    expect(initialCollabArtifact({ fixedFormat, initialHash: `#/collab?format=${differentFormat}&artifact=other-room` }, 'http://localhost:3100/')).toBe('')
  })
  it('preserves standalone room parsing and explicit initial deep links', () => {
    expect(initialCollabArtifact({}, 'http://localhost:3100/?artifact=current&format=docs#/collab')).toBe('current')
    expect(initialCollabArtifact({ initialHash: '#/collab?artifact=initial&format=docs' }, 'http://localhost:3100/?artifact=current')).toBe('initial')
  })
  it.each(formats)('accepts an explicit canonical %s room deep link without a redundant format query', (fixedFormat) => {
    expect(initialCollabArtifact({ fixedFormat, initialHash: `#/${fixedFormat}?feature=collab&artifact=canonical-room` }, 'http://localhost:3100/?format=pdf&artifact=unrelated'))
      .toBe('canonical-room')
    const differentFormat = fixedFormat === 'docs' ? 'sheets' : 'docs'
    expect(initialCollabArtifact({ fixedFormat, initialHash: `#/${differentFormat}?feature=collab&artifact=other-room` }, 'http://localhost:3100/'))
      .toBe('')
    expect(initialCollabArtifact({ fixedFormat, initialHash: `#/${fixedFormat}?feature=collab&format=${differentFormat}&artifact=canonical-room` }, 'http://localhost:3100/'))
      .toBe('canonical-room')
  })
  it.each(formats)('builds an isolated %s sharing link without altering other URL settings', (fixedFormat) => {
    const current = 'http://localhost:3100/demo?theme=dark&artifact=old&format=pdf#/docs?panel=history'
    const href = collabRoomHref(current, 'art a&b', fixedFormat)
    const url = new URL(href)
    expect(url.pathname).toBe('/demo')
    expect(url.searchParams.get('theme')).toBe('dark')
    expect(url.searchParams.has('artifact')).toBe(false)
    expect(url.searchParams.has('format')).toBe(false)
    expect(parseCollabQuery(href)).toEqual({ artifact: 'art a&b', format: fixedFormat })
    expect(url.hash).toMatch(/^#\/collab\?/)
    expect(collabRoomHref(current, 'anything')).toBe(current)
  })
  it('forwards fixed format and initial route through both collaboration runtime modes', () => {
    const page = source('./collabPage.tsx')
    expect(page).toContain('<CollabSimulator fixedFormat={fixedFormat} initialHash={initialHash} />')
    expect(page).toContain('<ServerCollabDemo fixedFormat={fixedFormat} initialHash={initialHash} />')
    expect(page).toContain('current room will close and unsaved demo edits will be lost')
    for (const file of ['./collabSimulator.tsx', './collabDemo.tsx']) {
      const implementation = source(file)
      expect(implementation).toContain('fixedFormat ?? selectedFormat')
      expect(implementation).toContain('fixedFormat === undefined && <SimFormatTabs')
      expect(implementation).toContain('if (fixedFormat !== undefined) return')
    }
  })
  it('suppresses every embedded collaboration URL-write effect', () => {
    for (const file of ['./collabDemo.tsx', './collab/roomChrome.tsx', './collab/pdf.tsx']) {
      expect(source(file)).toMatch(/useEffect\(\(\) => \{\s+if \(fixedFormat !== undefined\) return[\s\S]*?writeCollabQuery/)
    }
    for (const file of ['./collab/sheets.tsx', './collab/docs.tsx', './collab/slides.tsx', './collab/pdf.tsx']) {
      expect(source(file)).toContain('initialCollabArtifact(scope, window.location.href)')
      expect(source(file)).not.toContain('parseCollabQuery(window.location.href).artifact')
    }
  })
  it('starts embedded history in its locked content mode and preserves standalone switching', () => {
    const docs = renderToStaticMarkup(createElement(HistoryPage, { fixedFormat: 'docs' }))
    const sheets = renderToStaticMarkup(createElement(HistoryPage, { fixedFormat: 'sheets' }))
    const standalone = renderToStaticMarkup(createElement(HistoryPage))
    expect(docs).toContain('Revenue finished above target.')
    expect(docs).not.toContain('Name,Plan,Actual')
    expect(sheets).toContain('Name,Plan,Actual')
    expect(docs).not.toContain('aria-label="Diff mode"')
    expect(sheets).not.toContain('aria-label="Diff mode"')
    expect(standalone).toContain('aria-label="Diff mode"')
    expect(source('./pages/HistoryPage.tsx')).toContain("{ snapshot: BEFORE_TEXT, contentType: 'text/plain'")
  })
  it('gives repeated history, font, and shape panels unique referenced IDs', () => {
    const markup = renderToStaticMarkup(createElement(Fragment, null,
      createElement(HistoryPage, { fixedFormat: 'sheets' }), createElement(HistoryPage, { fixedFormat: 'docs' }),
      createElement(FontMetricsPage), createElement(FontMetricsPage), createElement(ShapesPage), createElement(ShapesPage),
      createElement(HistoryTimeline, { controller: {} as HistoryCommandController<unknown>, initialPage: { versions: [] } }),
      createElement(HistoryTimeline, { controller: {} as HistoryCommandController<unknown>, initialPage: { versions: [] } }),
    ))
    const ids = [...markup.matchAll(/\sid="([^"]+)"/g)].map((match) => match[1])
    expect(ids.length).toBeGreaterThan(10)
    expect(new Set(ids).size).toBe(ids.length)
    for (const match of markup.matchAll(/aria-labelledby="([^"]+)"/g)) for (const id of match[1].split(' ')) expect(ids).toContain(id)
  })
  it('scopes simulator and server-room heading IDs to their mounted instances', () => {
    for (const file of ['./collabSimulator.tsx', './collabSimChrome.tsx', './collab/roomChrome.tsx', './collab/pdf.tsx']) {
      expect(source(file)).toContain('useId()')
    }
    expect(source('./collabSimChrome.tsx')).toContain('aria-labelledby={titleId}')
    expect(source('./collab/roomChrome.tsx')).toContain('data-demo-busy={busy || metrics.pending > 0}')
    expect(source('./collab/pdf.tsx')).toContain('data-demo-busy={busy || metrics.pending > 0}')
  })
  it('renders multiple simulator frames and server room headings with distinct accessible IDs', () => {
    const metrics = { peers: 0, applied: 0, pending: 0 }
    const noop = () => undefined
    const room = { title: 'Room', connection: 'idle' as const, name: 'Mira', onNameChange: noop,
      artifact: '', onArtifactChange: noop, busy: false, onJoin: noop, status: 'Local only', error: '',
      proof: 'Protocol proof', editorTitle: 'Editor', editorHint: 'Sample', editorMeta: '', metrics,
      tryIt: null, children: null }
    const markup = renderToStaticMarkup(createElement(Fragment, null,
      ...formats.flatMap((format) => SIM_EDITORS.map((profile) => createElement(SimEditorFrame, {
        key: `${format}-${profile.id}`, profile, status: 'starting', presence: null, metrics, children: null,
      }))), createElement(CollabRoomChrome, room), createElement(CollabRoomChrome, room),
    ))
    const ids = [...markup.matchAll(/\sid="([^"]+)"/g)].map((match) => match[1])
    expect(ids).toHaveLength(10)
    expect(new Set(ids).size).toBe(ids.length)
    for (const match of markup.matchAll(/aria-labelledby="([^"]+)"/g)) expect(ids).toContain(match[1])
  })
})
