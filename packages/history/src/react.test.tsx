import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import type { HistoryCommandController } from './commands'
import {
  attachHistorySidebar,
  createHistorySidebarElement,
  defaultSize,
  HistoryTimeline,
  mountHistoryTimeline,
  trapHistoryDialogFocus,
} from './react'
import type { HistoryListPage } from './types'

const controller = {} as HistoryCommandController<unknown>
const page: HistoryListPage = {
  versions: [
    {
      id: 'v4', artifactId: 'book-1', sequence: 4, createdAt: 4_000, size: 1536,
      contentType: 'xlsx', author: { id: 'ada', kind: 'user', displayName: 'Ada' }, reason: 'save',
      description: 'Quarter close', retention: { policyId: 'audit', expiresAt: null, legalHold: true },
    },
    {
      id: 'v3', artifactId: 'book-1', sequence: 3, createdAt: 3_000, size: 512,
      contentType: 'xlsx', author: { id: 'agent-1', kind: 'agent' }, reason: 'restore', sourceVersionId: 'v1',
      retention: { policyId: 'standard', expiresAt: null, legalHold: false },
    },
  ],
  nextCursor: 'older',
}

describe('HistoryTimeline', () => {
  it('renders an accessible chronological ledger with host-gated actions', () => {
    const markup = renderToStaticMarkup(
      <HistoryTimeline
        controller={controller}
        initialPage={page}
        author={{ id: 'ada', kind: 'user' }}
        formatTime={(time) => `time-${time}`}
      />,
    )
    expect(markup).toContain('aria-label="Saved versions"')
    expect(markup).toContain('Ada')
    expect(markup).toContain('Quarter close')
    expect(markup).toContain('Restored from v1')
    expect(markup).toContain('Save version')
    expect(markup).toContain('Load older versions')
    expect(markup.match(/>Restore</g)).toHaveLength(1)
    expect(markup).toContain('time-4000')
  })

  it('removes capture/restore controls when the host denies them', () => {
    const markup = renderToStaticMarkup(
      <HistoryTimeline controller={controller} initialPage={page} canCapture={false} canRestore={false} />,
    )
    expect(markup).not.toContain('Save version')
    expect(markup).not.toContain('>Restore<')
    expect(markup).toContain('Browse saved states without changing the open workbook.')
  })

  it('formats bounded byte sizes for timeline metadata', () => {
    expect(defaultSize(100)).toBe('100 B')
    expect(defaultSize(1536)).toBe('1.5 KB')
    expect(defaultSize(12 * 1024)).toBe('12 KB')
    expect(defaultSize(2.5 * 1024 ** 2)).toBe('2.5 MB')
  })

  it('creates a sidebar host and mounts the timeline through an injected root', () => {
    const doc = {
      createElement: (tag: string) => {
        const element = { tagName: tag.toUpperCase(), className: '', attributes: {} as Record<string, string> }
        return Object.assign(element, {
          setAttribute(name: string, value: string) { element.attributes[name] = value },
        })
      },
    }
    const aside = createHistorySidebarElement(doc as unknown as Pick<Document, 'createElement'>)
    expect(aside.className).toBe('ioc-history-sidebar')
    expect((aside as unknown as { attributes: Record<string, string> }).attributes['data-ioc-history-sidebar']).toBe('true')

    const rendered: unknown[] = []
    const unmount = vi.fn()
    const parent = { children: [] as unknown[], appendChild(node: unknown) { this.children.push(node); return node }, removeChild(node: unknown) { this.children = this.children.filter((child) => child !== node); return node } }
    const attached = attachHistorySidebar(
      parent as unknown as Element,
      { controller, initialPage: page },
      () => ({ render: (node) => { rendered.push(node) }, unmount }),
      doc as unknown as Pick<Document, 'createElement'>,
    )
    expect(parent.children).toHaveLength(1)
    expect(rendered).toHaveLength(1)
    attached.unmount()
    expect(unmount).toHaveBeenCalledOnce()
    expect(parent.children).toHaveLength(0)

    const mounted = mountHistoryTimeline(
      aside as unknown as Element,
      { controller, initialPage: page },
      () => ({ render: (node) => { rendered.push(node) }, unmount }),
    )
    mounted.unmount()
    expect(unmount).toHaveBeenCalledTimes(2)
  })

  it('traps restore-dialog focus and cancels on Escape', () => {
    const cancel = vi.fn()
    const preventDefault = vi.fn()
    const first = { focus: vi.fn() }
    const last = { focus: vi.fn() }
    const root = { querySelectorAll: () => [first, last] }
    trapHistoryDialogFocus({ key: 'Escape', shiftKey: false, preventDefault }, root as never, cancel, first as never)
    expect(cancel).toHaveBeenCalledOnce()
    trapHistoryDialogFocus({ key: 'Tab', shiftKey: false, preventDefault }, root as never, cancel, last as never)
    expect(first.focus).toHaveBeenCalledOnce()
    trapHistoryDialogFocus({ key: 'Tab', shiftKey: true, preventDefault }, root as never, cancel, first as never)
    expect(last.focus).toHaveBeenCalledOnce()
  })
})
