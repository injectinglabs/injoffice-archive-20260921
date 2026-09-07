import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const source = (name: string) => readFileSync(new URL(`./${name}`, import.meta.url), 'utf8')

describe('collaboration playground', () => {
  it('defaults to a two-editor browser simulation and preserves the server integration mode', () => {
    const page = source('collabPage.tsx')
    const simulator = source('collabSimulator.tsx')

    expect(page).toContain("useState<Mode>('simulation')")
    expect(page).toContain('Two-editor simulation')
    expect(page).toContain('HTTP + SSE integration')
    expect(simulator.match(/<SimulatedEditor/g)).toHaveLength(1)
    expect(simulator).toContain('EDITORS.map')
    expect(simulator).toContain('independent Univer instances')
    expect(simulator).toContain('<SimFormatTabs')
    expect(source('collabSimChrome.tsx')).toContain('collab-format-bar')
    expect(source('collabDemo.tsx')).toContain('<SimFormatTabs')
    expect(simulator).toContain('<CollabSimulatorDocs')
    expect(simulator).toContain('<CollabSimulatorSlides')
    expect(simulator).toContain('<CollabSimulatorPdf')
  })

  it('wires docs, slides, and PDF two-pane simulators to the in-memory hub', () => {
    const docs = source('collabSimulatorDocs.tsx')
    const slides = source('collabSimulatorSlides.tsx')
    const pdf = source('collabSimulatorPdf.tsx')
    expect(docs).toContain('DocPresenceManager')
    expect(docs).toContain('hub.connect<DocSelection>')
    expect(docs).toContain('receiveTransaction')
    expect(docs).not.toContain('createHttpCollabTransport')
    expect(slides).toContain('DeckPresenceManager')
    expect(slides).toContain('hub.connect<DeckSelection>')
    expect(slides).toContain('seedCollabDeck')
    expect(slides).not.toContain('createHttpCollabTransport')
    expect(pdf).toContain('PdfPresenceManager')
    expect(pdf).toContain('hub.connect<PdfSelection>')
    expect(pdf).toContain('encodePdfCollabOperation')
    expect(pdf).toContain('makeCollabPdfSample')
    expect(pdf).not.toContain('createHttpCollabTransport')
  })

  it('states the browser-only boundary without presenting it as file persistence', () => {
    const simulator = source('collabSimulator.tsx')
    expect(simulator).toContain('Memory only')
    expect(simulator).toContain('No file is uploaded')
    expect(simulator).toContain('refreshing clears the session')
  })

  it('boots and disposes nested Univer roots outside the Strict Mode effect cycle', () => {
    const simulator = source('collabSimulator.tsx')
    const serverSheets = source('collab/sheets.tsx')
    expect(simulator).toContain("window.setTimeout(startEditor, profile.id === 'noah' ? 75 : 0)")
    expect(simulator).toContain('window.clearTimeout(startTimer)')
    expect(simulator).toContain('window.setTimeout(() => univer.dispose(), 0)')
    expect(serverSheets).toContain('useEffect(() => deferNestedReactRootStart(() => {')
    expect(serverSheets).toContain('if (!editorRef.current) return')
    expect(serverSheets).toContain('window.setTimeout(() => univer.dispose(), 0)')
  })

  it('projects remote drafts into the peer cell without enabling Univer popups', () => {
    const simulator = source('collabSimulator.tsx')
    const ghost = source('cellGhost.tsx')
    expect(simulator).toContain('<RemoteCellGhost')
    expect(simulator).toContain('peerLabelComponent: false')
    expect(simulator).toMatch(/uncommitted draft appears inside that cell/i)
    expect(ghost).toContain('measureUniverCellAnchor')
    expect(ghost).toContain('cellGhostViewModel')
    expect(ghost).not.toContain('attachRangePopup')
  })
})
