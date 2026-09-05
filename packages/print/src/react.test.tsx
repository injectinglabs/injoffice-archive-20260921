import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { PrintConfigurationCommandController } from './commands'
import { PrintManager } from './manager'
import { PrintPreviewManager, type PrintPreviewDocument } from './preview'
import { canonicalPageRange, PrintWorkspace } from './react'

function fixture() {
  const manager = new PrintManager({ print: vi.fn() }, 'sheet-1')
  const controller = new PrintConfigurationCommandController(manager, { save: vi.fn() })
  const previewManager = new PrintPreviewManager<{ label: string }>({ renderPreview: () => ({ totalPages: 0, pages: [] }) })
  const initialDocument: PrintPreviewDocument<{ label: string }> = {
    sessionId: 'server-preview',
    snapshot: manager.snapshot(),
    selection: { kind: 'all' },
    totalPages: 2,
    pages: [
      { number: 1, widthPoints: 595, heightPoints: 842, payload: { label: 'Rendered sales table' } },
      { number: 2, widthPoints: 595, heightPoints: 842, payload: { label: 'Rendered totals' } },
    ],
  }
  return { controller, previewManager, initialDocument }
}

describe('PrintWorkspace', () => {
  it('renders an accessible settings workflow and host-provided pages', () => {
    const props = fixture()
    const markup = renderToStaticMarkup(<PrintWorkspace {...props} renderPage={(page) => <div>{page.payload.label}</div>} />)

    expect(markup).toContain('role="dialog"')
    expect(markup).toContain('aria-modal="true"')
    expect(markup).toContain('aria-label="Print settings"')
    expect(markup).toContain('aria-label="Print preview"')
    expect(markup).toContain('aria-label="Preview pages"')
    expect(markup).toContain('Rendered sales table')
    expect(markup).toContain('Rendered totals')
    expect(markup).toContain('Page 2')
    expect(markup).toContain('Show row and column headings')
    expect(markup).toContain('Repeat on every page')
    expect(markup).toContain('Refresh preview')
  })

  it('lets hosts remove controls, preview, and actions independently', () => {
    const props = fixture()
    const markup = renderToStaticMarkup(<PrintWorkspace {...props} renderPage={() => null} config={{
      modal: false,
      showPreview: false,
      showPrintAction: false,
      showCloseAction: false,
      controls: { paper: false, scale: false, margins: false, repeatedTitles: false, content: false, alignment: false, headerFooter: false, watermark: false },
    }} />)

    expect(markup).toContain('role="region"')
    expect(markup).not.toContain('aria-modal')
    expect(markup).not.toContain('Print preview')
    expect(markup).not.toContain('Paper size')
    expect(markup).not.toContain('>Print</button>')
    expect(markup).not.toContain('>Close</button>')
    expect(markup).toContain('Print area')
  })

  it('validates one-based preview ranges before starting a session', () => {
    expect(canonicalPageRange(2, 4, 5)).toEqual({ kind: 'range', startPage: 2, endPage: 4 })
    expect(canonicalPageRange(0, 1, 5)).toBeNull()
    expect(canonicalPageRange(4, 2, 5)).toBeNull()
    expect(canonicalPageRange(2, 6, 5)).toBeNull()
  })
})
