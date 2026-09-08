import { useEffect, useState } from 'react'
import { DsSegment } from '../design-system/primitives'
import '../design-system/live-create-edit.css'
import UniverEditor from '../UniverEditor'
import { parseSheetsView, parseSurface, type SheetsView } from '../route'
import NativeRoundTripPage from './NativeRoundTripPage'
import SheetsToolsPage from './SheetsToolsPage'

const SHEETS_VIEWS: readonly { id: SheetsView; label: string }[] = [
  { id: 'editor', label: 'Edit workbook' },
  { id: 'native', label: 'Test XLSX round trip' },
  { id: 'tools', label: 'Package tools' },
]

export default function SheetsPage({ initialHash }: { initialHash?: string } = {}) {
  const [view, setView] = useState<SheetsView>(() => parseSheetsView(initialHash))
  useEffect(() => {
    // Retained scroll sections must not reset when another section owns the URL.
    const syncView = () => {
      if (parseSurface() === 'sheets') setView(parseSheetsView())
    }
    window.addEventListener('hashchange', syncView)
    return () => window.removeEventListener('hashchange', syncView)
  }, [])
  const selectView = (next: SheetsView) => {
    setView(next)
    window.location.hash = `#/sheets?view=${next}`
  }
  return (
    <div className="tool-page tool-page--flush ds">
      <div className="view-switcher ds-workstrip" role="group" aria-label="Spreadsheet demonstration">
        <DsSegment
          label="Spreadsheet demonstration"
          value={view}
          onChange={(id) => selectView(id as SheetsView)}
          options={SHEETS_VIEWS}
        />
        <span className="ds-muted">
          {view === 'editor' ? 'Runs in your browser' : view === 'native' ? 'Original bytes stay in this browser' : 'Sparklines, print, outlines, and exchange'}
        </span>
      </div>
      <div className="tool-page-fill">{view === 'editor' ? <UniverEditor /> : view === 'native' ? <NativeRoundTripPage /> : <SheetsToolsPage />}</div>
    </div>
  )
}
