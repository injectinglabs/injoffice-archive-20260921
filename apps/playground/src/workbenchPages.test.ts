import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const pageSource = (name: string) => readFileSync(new URL(`./pages/${name}.tsx`, import.meta.url), 'utf8')

describe('playground workbench page contract', () => {
  it('keeps analytical demos in the shared tool workbench', () => {
    for (const page of ['ChartsPage', 'ConnectorsPage', 'PivotsPage', 'ShapesPage', 'FormulasPage', 'HistoryPage', 'AgentPage']) {
      const source = pageSource(page)
      expect(source, page).toMatch(/className="[^"]*\btool-page\b[^"]*"/)
      expect(source, page).toMatch(/role="(?:toolbar|search)"/)
      expect(source, page).toContain('aria-label=')
    }
  })

  it('keeps standalone actions in the shared button hierarchy', () => {
    for (const page of ['ChartsPage', 'ConnectorsPage', 'HistoryPage']) {
      expect(pageSource(page), page).toContain('className="workbench-button"')
    }

    expect(pageSource('PptxNativePage')).toContain('workbench-button workbench-button--primary')
  })

  it('announces capability runtime and changing result states', () => {
    for (const page of ['PptxAuthoredPage', 'PptxRenderPage', 'FontMetricsPage']) {
      expect(pageSource(page), page).toMatch(/className="capability-runtime" role="status"/)
    }

    expect(pageSource('PptxNativePage')).toContain('aria-label="PPTX processing runtime"')
    expect(pageSource('PptxNativePage')).toContain('aria-live="polite"')
    expect(readFileSync(new URL('./App.tsx', import.meta.url), 'utf8')).toContain('role="status"')
    expect(pageSource('FormulasPage')).toContain('aria-live="polite"')
  })

  it('wires the PDF workbench to browser-safe package APIs and a Vite-only Node host', () => {
    const page = pageSource('PdfPage')
    expect(page).not.toContain('pdfNodeHost')
    expect(page).not.toContain('formRedact')
    expect(page).not.toContain('textEdit')
    expect(page).not.toContain('renderPageToPng')
    expect(page).toContain('applyPdfPageOp')
    expect(page).toContain('applyPdfMarkup')
    expect(page).toContain('applyPdfPlacedDrawing')
    expect(page).toContain('applyPdfFormValues')
    expect(page).toContain('pdfFormResultMessage(result)')
    expect(page).toContain('if (result.applied > 0) await applyBytes')
    expect(page).toContain('{formNotice && <p role="status"')
    expect(page).not.toContain("'Applied form values.'")
    expect(page).toContain('applyPdfStamp')
    expect(page).toContain("type: 'nUp', n: 4")
    expect(page).toContain('images/insert')
    expect(page).toContain('redact/form')
    expect(page).toContain('PDF_NODE_PREFIX')

    const vite = readFileSync(new URL('../vite.config.ts', import.meta.url), 'utf8')
    expect(vite).toContain("name: 'injoffice-pdf-node-host'")
    expect(vite).toContain("apply: 'serve'")
    expect(vite).toContain("./pdfNodeHost.ts")
    expect(vite).not.toContain('/v1/pdf')
  })

  it('clears the active PPTX session before replacement I/O can fail', () => {
    const source = pageSource('PptxNativePage')
    const extract = source.slice(source.indexOf('const extractFile'), source.indexOf('const loadSample'))
    const sample = source.slice(source.indexOf('const loadSample'), source.indexOf('const chooseTarget'))

    expect(extract.indexOf('clearSession()')).toBeGreaterThanOrEqual(0)
    expect(extract.indexOf('clearSession()')).toBeLessThan(extract.indexOf('try {'))
    expect(sample.indexOf('clearSession()')).toBeGreaterThanOrEqual(0)
    expect(sample.indexOf('clearSession()')).toBeLessThan(sample.indexOf('try {'))
  })
})
