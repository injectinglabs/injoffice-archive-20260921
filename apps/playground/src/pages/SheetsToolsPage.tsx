import { useMemo, useState } from 'react'
import { PrintWorkspace } from '../../../../packages/print/src/react'
import { DsButton, DsField, DsSelect } from '../design-system/primitives'
import '../design-system/live-create-edit.css'
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
    <section className="tool-page ds" data-demo-surface="sheets-tools" aria-label="Spreadsheet package tools">
      <div className="tool-page__controls ds-workstrip" role="toolbar" aria-label="Sheet tool proofs">
        <span>Sparklines, print, outlines, and exchange</span>
        <span className="tool-page__status ds-muted">{printNote}</span>
      </div>

      <div className="tool-page__grid ds-cards">
        <section className="tool-card ds-panel" aria-labelledby="sparkline-title">
          <span className="tool-eyebrow ds-eyebrow">@injoffice/sparklines</span>
          <div className="tool-card__heading"><div><h2 id="sparkline-title">Sparkline</h2></div><span className="tool-chip">{sparkType}</span></div>
          <DsField label="Sparkline">
            <DsSelect value={sparkType} aria-label="Sparkline type" onChange={(event) => setSparkType(event.target.value as SparklineType)}>
              <option value="line">line</option>
              <option value="column">column</option>
              <option value="win-loss">win-loss</option>
            </DsSelect>
          </DsField>
          <img className="ds-spark" alt="Sparkline" src={`data:image/svg+xml;utf8,${encodeURIComponent(sparkline.svg)}`} width={180} height={48} />
          <p className="tool-note ds-muted">{SAMPLE_SPARKLINE_VALUES.join(', ')}</p>
        </section>
        <section className="tool-card ds-panel" aria-labelledby="print-title">
          <span className="tool-eyebrow ds-eyebrow">@injoffice/print</span>
          <strong id="print-title">Print workspace</strong>
          <p className="tool-note ds-muted">{printNote}</p>
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
        <section className="tool-card ds-panel" aria-labelledby="outline-title">
          <span className="tool-eyebrow ds-eyebrow">@injoffice/outlines</span>
          <strong id="outline-title">Row groups</strong>
          <div className="ds-tools-grid">
            {ROW_LABELS.map((label, index) => (
              <div key={label} style={{ display: 'contents' }}>
                <span>{index === 1 ? (visible[2] ? '▾' : '▸') : ''}</span>
                {visible[index] ? <span>{label}</span> : <s>{label}</s>}
              </div>
            ))}
          </div>
          <DsButton className="workbench-button" onClick={() => {
            const next = outlines.manager.toggle('q1')
            if (next.ok) setHidden([...outlines.hidden])
          }}>{visible[2] ? 'Collapse Q1 rows' : 'Expand Q1 rows'}</DsButton>
        </section>
        <section className="tool-card ds-panel" aria-labelledby="exchange-title">
          <span className="tool-eyebrow ds-eyebrow">@injoffice/xlsx-exchange</span>
          <strong id="exchange-title">Exchange job</strong>
          <p className="tool-note ds-muted">{exchangeNote}</p>
          <DsButton className="workbench-button" onClick={() => {
            const { manager } = createPlaygroundExchange()
            const job = manager.importSnapshot({ source: { kind: 'bytes', bytes: new Uint8Array([80, 75, 3, 4]), name: 'demo.xlsx' } })
            void job.result.then((result) => setExchangeNote(`${job.id} imported ${result.snapshot.name} (${result.snapshot.bytes} bytes).`))
          }}>Import XLSX job</DsButton>
        </section>
      </div>
    </section>
  )
}
