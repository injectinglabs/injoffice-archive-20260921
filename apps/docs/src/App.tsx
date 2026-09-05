import { useEffect, useState, type ReactNode } from 'react'
import { Layout } from './components/Layout'
import { parsePath, parseSection, navigate, isDocsHash } from './route'
import type { DocId } from './catalog'
import { HomePage } from './pages/HomePage'
import { ConceptsPage, InstallationPage, IntroductionPage, QuickstartPage } from './pages/GettingStartedPages'
import { DocxPage, PptxPage, XlsxPage } from './pages/NativePages'
import {
  ChartsPage,
  CollabPage,
  ConnectorsPage,
  FontMetricsPage,
  FormulasPage,
  HistoryPage,
  PdfGuidePage,
  PivotsPage,
  ShapesPage,
  SheetsEditorPage,
  SlidesGuidePage,
} from './pages/FeaturePages'
import { ComparePage } from './pages/ComparePage'
import { ShowcasePage } from './pages/ShowcasePage'
import { ReferencePage } from './pages/ReferencePage'
import './styles.css'

const PAGES: Record<DocId, () => ReactNode> = {
  home: () => <HomePage />,
  introduction: () => <IntroductionPage />,
  installation: () => <InstallationPage />,
  concepts: () => <ConceptsPage />,
  quickstart: () => <QuickstartPage />,
  xlsx: () => <XlsxPage />,
  docx: () => <DocxPage />,
  pptx: () => <PptxPage />,
  'sheets-editor': () => <SheetsEditorPage />,
  charts: () => <ChartsPage />,
  pivots: () => <PivotsPage />,
  shapes: () => <ShapesPage />,
  connectors: () => <ConnectorsPage />,
  formulas: () => <FormulasPage />,
  collab: () => <CollabPage />,
  history: () => <HistoryPage />,
  'font-metrics': () => <FontMetricsPage />,
  pdf: () => <PdfGuidePage />,
  slides: () => <SlidesGuidePage />,
  compare: () => <ComparePage />,
  showcase: () => <ShowcasePage />,
  reference: () => <ReferencePage />,
}

export default function App() {
  const [id, setId] = useState<DocId>(() => parsePath())

  useEffect(() => {
    const sync = () => setId(parsePath())
    window.addEventListener('hashchange', sync)
    const onClick = (event: MouseEvent) => {
      const anchor = (event.target as HTMLElement | null)?.closest('a')
      if (!anchor || event.defaultPrevented || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return
      if (anchor.target === '_blank') return
      const href = anchor.getAttribute('href')
      if (!href?.startsWith('#/guides')) return
      event.preventDefault()
      navigate(href)
    }
    document.addEventListener('click', onClick)
    return () => {
      window.removeEventListener('hashchange', sync)
      document.removeEventListener('click', onClick)
    }
  }, [])

  useEffect(() => {
    document.title = id === 'home' ? 'InjOffice Docs' : `${document.querySelector('.inj-docs h1')?.textContent ?? 'InjOffice'} · InjOffice Docs`
    const section = parseSection()
    if (section) document.getElementById(section)?.scrollIntoView({ block: 'start' })
  }, [id])

  if (!isDocsHash()) return null
  const render = PAGES[id]
  return <Layout id={id}>{render()}</Layout>
}
