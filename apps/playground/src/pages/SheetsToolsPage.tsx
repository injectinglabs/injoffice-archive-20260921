import { useMemo, useState } from 'react'
import { PrintWorkspace } from '../../../../packages/print/src/react'
import {
  SAMPLE_SPARKLINE_VALUES,
  createPlaygroundExchange,
  createPlaygroundOutlines,
  createPlaygroundPrintWorkspace,
  createPlaygroundSparkline,
  visibleOutlineRows,
} from '../sheetTools'
import type { SparklineType } from '../../../../packages/sparklines/src/types'

const ROW_LABELS = ['Header', 'North', 'Q1', 'Q2', 'Q3', 'South']

export default function SheetsToolsPage() {
  const [sparkType, setSparkType] = useState<SparklineType>('line')
  const sparkline = useMemo(() => createPlaygroundSparkline(sparkType, SAMPLE_SPARKLINE_VALUES), [sparkType])
  const print = useMemo(() => createPlaygroundPrintWorkspace(), [])
  const [printNote, setPrintNote] = useState('Print workspace is mounted.')
  const outlines = useMemo(() => createPlaygroundOutlines(), [])
  const [hidden, setHidden] = useState(() => {
    outlines.manager.add({ id: 'q1', sheetId: 'sheet-1', axis: 'row', start: 2, end: 4, collapsed: true })
    return [...outlines.hidden]
  })
  const visible = visibleOutlineRows(ROW_LABELS.length, hidden)
  const [exchangeNote, setExchangeNote] = useState('No exchange jobs yet.')

  return (
    <section className="tool-page" data-demo-surface="sheets-tools" aria-label="Spreadsheet package tools">
      <div className="tool-page__controls" role="toolbar" aria-label="Sheet tool proofs">
        <label className="tool-field">
          Sparkline
          <select value={sparkType} aria-label="Sparkline type" onChange={(event) => setSparkType(event.target.value as SparklineType)}>
            <option value="line">line</option>
            <option value="column">column</option>
            <option value="win-loss">win-loss</option>
          </select>
        </label>
        <span className="tool-page__status">{printNote}</span>
        <button className="workbench-button" type="button" onClick={() => {
          const next = outlines.manager.toggle('q1')
          if (next.ok) setHidden([...outlines.hidden])
        }}>{visible[2] ? 'Collapse Q1 rows' : 'Expand Q1 rows'}</button>
        <button className="workbench-button" type="button" onClick={() => {
          const { manager } = createPlaygroundExchange()
          const job = manager.importSnapshot({ source: { kind: 'bytes', bytes: new Uint8Array([80, 75, 3, 4]), name: 'demo.xlsx' } })
          void job.result.then((result) => setExchangeNote(`${job.id} imported ${result.snapshot.name} (${result.snapshot.bytes} bytes).`))
        }}>Import XLSX job</button>
        <span className="tool-page__status">Sparklines, print, outlines, and exchange stay on this existing spreadsheet demo</span>
      </div>

      <div className="tool-page__grid">
        <section className="tool-card" aria-labelledby="sparkline-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">@injoffice/sparklines</span><h2 id="sparkline-title">Sparkline</h2></div><span className="tool-chip">{sparkType}</span></div>
          <img alt="Sparkline" src={`data:image/svg+xml;utf8,${encodeURIComponent(sparkline.svg)}`} width={180} height={48} />
          <p className="tool-note">{SAMPLE_SPARKLINE_VALUES.join(', ')}</p>
        </section>
        <section className="tool-card" aria-labelledby="print-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">@injoffice/print</span><h2 id="print-title">Print workspace</h2></div></div>
          <PrintWorkspace
            controller={print.controller}
            previewManager={print.previewManager}
            renderPage={(page) => <div>{page.payload.label}</div>}
            config={{ modal: false, title: 'Print preview' }}
            onPrinted={() => {
              const last = print.printed.at(-1)
              setPrintNote(last ? `Host printed ${last.direction} ${last.paperSize}.` : 'Host printed the current layout.')
            }}
          />
        </section>
        <section className="tool-card" aria-labelledby="outline-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">@injoffice/outlines</span><h2 id="outline-title">Row groups</h2></div></div>
          <ol className="formula-list">
            {ROW_LABELS.map((label, index) => (
              <li key={label} className={visible[index] ? undefined : 'tool-empty'}>{visible[index] ? label : `${label} (hidden)`}</li>
            ))}
          </ol>
        </section>
        <section className="tool-card" aria-labelledby="exchange-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">@injoffice/xlsx-exchange</span><h2 id="exchange-title">Exchange job</h2></div></div>
          <p className="tool-note">{exchangeNote}</p>
        </section>
      </div>
    </section>
  )
}
