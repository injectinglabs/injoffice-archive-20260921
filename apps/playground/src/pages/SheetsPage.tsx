import { useState } from 'react'
import UniverEditor from '../UniverEditor'
import { parseSheetsView, type SheetsView } from '../route'
import NativeRoundTripPage from './NativeRoundTripPage'
import SheetsToolsPage from './SheetsToolsPage'

export default function SheetsPage() {
  const [view, setView] = useState<SheetsView>(() => parseSheetsView())
  return (
    <div className="tool-page tool-page--flush">
      <div className="view-switcher" role="group" aria-label="Spreadsheet demonstration">
        <div className="tool-segment">
          <button type="button" aria-pressed={view === 'editor'} onClick={() => setView('editor')}>Edit workbook</button>
          <button type="button" aria-pressed={view === 'native'} onClick={() => setView('native')}>Test XLSX round trip</button>
          <button type="button" aria-pressed={view === 'tools'} onClick={() => setView('tools')}>Package tools</button>
        </div>
        <span>{view === 'editor' ? 'Runs in your browser' : view === 'native' ? 'Runs locally by default · server fallback is explicit' : 'Sparklines, print, outlines, and exchange'}</span>
      </div>
      <div className="tool-page-fill">{view === 'editor' ? <UniverEditor /> : view === 'native' ? <NativeRoundTripPage /> : <SheetsToolsPage />}</div>
    </div>
  )
}
