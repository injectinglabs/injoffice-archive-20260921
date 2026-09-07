import { useEffect, useState } from 'react'
import UniverEditor from '../UniverEditor'
import { parseSheetsView, type SheetsView } from '../route'
import NativeRoundTripPage from './NativeRoundTripPage'
import SheetsToolsPage from './SheetsToolsPage'

export default function SheetsPage() {
  const [view, setView] = useState<SheetsView>(() => parseSheetsView())
  useEffect(() => {
    const syncView = () => setView(parseSheetsView())
    window.addEventListener('hashchange', syncView)
    return () => window.removeEventListener('hashchange', syncView)
  }, [])
  const selectView = (next: SheetsView) => {
    setView(next)
    window.location.hash = `#/sheets?view=${next}`
  }
  return (
    <div className="tool-page tool-page--flush">
      <div className="view-switcher" role="group" aria-label="Spreadsheet demonstration">
        <div className="tool-segment">
          <button type="button" aria-pressed={view === 'editor'} onClick={() => selectView('editor')}>Edit workbook</button>
          <button type="button" aria-pressed={view === 'native'} onClick={() => selectView('native')}>Test XLSX round trip</button>
          <button type="button" aria-pressed={view === 'tools'} onClick={() => selectView('tools')}>Package tools</button>
        </div>
        <span>{view === 'editor' ? 'Runs in your browser' : view === 'native' ? 'Runs locally by default · server fallback is explicit' : 'Sparklines, print, outlines, and exchange'}</span>
      </div>
      <div className="tool-page-fill">{view === 'editor' ? <UniverEditor /> : view === 'native' ? <NativeRoundTripPage /> : <SheetsToolsPage />}</div>
    </div>
  )
}
