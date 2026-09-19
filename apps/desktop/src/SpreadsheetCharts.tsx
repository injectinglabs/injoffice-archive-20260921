import { useState } from 'react';
import type { XlsxNativeChart } from '@injoffice/xlsx-wasm';
import type { RangeRef } from '@injoffice/sheets/browser';
import type { SheetOperation } from './spreadsheetCommands';
import { address } from './spreadsheetCommands';
import { chartAnchor, chartGeometry, chartNumber, chartSelectionReason, CHART_COLORS } from './spreadsheet-chart-geometry';
import './spreadsheet-charts.css';

export function SpreadsheetChartPreview({ chart }: { chart: XlsxNativeChart }) {
  const g = chartGeometry(chart), horizontal = chart.chart_type === 'bar', n = chart.categories.length, count = chart.series.length;
  const left = horizontal ? 86 : 66, top = 12, width = horizontal ? 322 : 342, height = 210;
  const labelStep = Math.max(1, Math.ceil(n / (horizontal ? 8 : 6)));
  const short = (text: string) => text.length > 13 ? `${text.slice(0, 12)}…` : text;
  return <svg className="sheet-chart-preview" viewBox="0 0 430 258" role="img" aria-label={`${chart.title || 'Chart'}: ${chart.chart_type}, ${n} categories, ${count} series. Full values follow in Chart data.`}>
    {g.ticks.map((tick, i) => horizontal ? <g key={i}><line x1={left + tick.position * width} x2={left + tick.position * width} y1={top} y2={top + height} stroke="#e3e9e6"/><text x={left + tick.position * width} y={top + height + 19} textAnchor="middle">{chartNumber(tick.value)}</text></g> : <g key={i}><line x1={left} x2={left + width} y1={top + (1 - tick.position) * height} y2={top + (1 - tick.position) * height} stroke="#e3e9e6"/><text x={left - 8} y={top + (1 - tick.position) * height + 4} textAnchor="end">{chartNumber(tick.value)}</text></g>)}
    {chart.categories.map((label, i) => i % labelStep === 0 ? <text key={i} x={horizontal ? left - 7 : left + (i + .5) * width / n} y={horizontal ? top + (i + .5) * height / n + 4 : top + height + 20} textAnchor={horizontal ? 'end' : 'middle'}><title>{label}</title>{short(label)}</text> : null)}
    {g.normalized.map((values, si) => chart.chart_type === 'line' ? <path key={si} fill="none" stroke={CHART_COLORS[si]} strokeWidth="2" d={values.map((v, i) => `${i ? 'L' : 'M'}${left + (i + .5) * width / n},${top + (1 - g.position(v)) * height}`).join(' ')}/> : <g key={si} fill={CHART_COLORS[si]}>{values.map((v, i) => {
      const slot = (horizontal ? height : width) / n, bar = slot * .75 / count, p = g.position(v), z = g.zero;
      return horizontal ? <rect key={i} x={left + Math.min(z, p) * width} y={top + (i + .125) * slot + si * bar} width={Math.abs(p - z) * width} height={bar}><title>{`${chart.categories[i]} — ${chart.series[si]!.name}: ${chart.series[si]!.values[i]}`}</title></rect> : <rect key={i} x={left + (i + .125) * slot + si * bar} y={top + (1 - Math.max(z, p)) * height} width={bar} height={Math.abs(p - z) * height}><title>{`${chart.categories[i]} — ${chart.series[si]!.name}: ${chart.series[si]!.values[i]}`}</title></rect>;
    })}</g>)}
    {chart.chart_type === 'line' && n === 1 && g.normalized.map((values, i) => <circle key={i} cx={left + width / 2} cy={top + (1 - g.position(values[0]!)) * height} r="3" fill={CHART_COLORS[i]}/>)}
    {horizontal ? <line x1={left + g.zero * width} x2={left + g.zero * width} y1={top} y2={top + height} stroke="#77867e"/> : <line x1={left} x2={left + width} y1={top + (1 - g.zero) * height} y2={top + (1 - g.zero) * height} stroke="#77867e"/>}
  </svg>;
}
function ChartTitleForm({chart, disabled, onExecute}: {chart:XlsxNativeChart; disabled:boolean; onExecute(operations:SheetOperation[],message:string):Promise<boolean>}) {
  const [title,setTitle] = useState(chart.title);
  return <form className="sheet-chart-title" onSubmit={event=>{event.preventDefault();if(disabled || title===chart.title)return;void onExecute([{kind:'chart.update',identity:chart.identity,expected_fingerprint_sha256:chart.fingerprint_sha256,chart_type:chart.chart_type as 'column'|'bar'|'line',title,range:chart.range,anchor:chart.anchor}], 'Chart title updated');}}>
    <label>Chart title<input aria-label="Chart title" maxLength={1024} disabled={disabled} value={title} onChange={event=>setTitle(event.target.value)}/></label><button type="submit" disabled={disabled || title===chart.title}>Apply title</button>
    {title!==chart.title && <small>Apply to save this title in the workbook.</small>}
  </form>;
}
export function SpreadsheetCharts({ charts, sheetId, range, disabled, error, onExecute, onClose }: {
  charts: XlsxNativeChart[]; sheetId: string; range: RangeRef; disabled: boolean; error?: string;
  onExecute(operations: SheetOperation[], message: string): Promise<boolean>; onClose(): void;
}) {
  const [type, setType] = useState<'column'|'bar'|'line'>('column');
  const [newTitle,setNewTitle] = useState('');
  const reason = chartSelectionReason(range), local = charts.filter(chart => chart.sheet_id === sheetId || !chart.sheet_id);
  const options = <><option value="column">Column</option><option value="bar">Bar</option><option value="line">Line</option></>;
  return <aside className="sheet-charts" aria-label="Worksheet charts">
    <header><h2>Charts</h2><button onClick={onClose} aria-label="Close charts">Close</button></header>
    <div className="sheet-chart-create"><label>New chart title<input aria-label="New chart title" maxLength={1024} value={newTitle} onChange={event=>setNewTitle(event.target.value)}/></label><label>New chart type<select aria-label="New chart type" value={type} onChange={event => setType(event.target.value as typeof type)}>{options}</select></label><button disabled={disabled || !!reason || !!error || charts.length >= 32} title={reason} onClick={() => void onExecute([{kind:'chart.insert',chart_type:type,title:newTitle,range,anchor:chartAnchor(range)}], 'Chart inserted into workbook')}>Insert selected range</button><p>{reason ?? `Source ${address(range)}:${address({row:range.end_row,column:range.end_column})}. First row: series names. First column: categories.`}</p><small>Literal numbers only. Formulas, blanks, merged or hidden data are unsupported. Charts save inside the XLSX file.</small></div>
    {error && <p className="sheet-chart-refusal" role="status">Charts are unavailable: {error}. Cell editing remains available.</p>}
    {!local.length && !error && <p className="sheet-chart-empty">No charts on this sheet. Select labels and numeric values in the grid, then insert a chart.</p>}
    {local.map((chart, index) => <section className="sheet-chart" key={`${chart.identity.part}:${index}`} aria-label={chart.title || `Chart ${index + 1}`}>
      <h3>{chart.title || `Chart ${index + 1}`}</h3>
      {chart.editable ? <>
        <ChartTitleForm key={chart.fingerprint_sha256} chart={chart} disabled={disabled} onExecute={onExecute}/>
        <SpreadsheetChartPreview chart={chart}/>
        <ul className="sheet-chart-legend">{chart.series.map((series, i) => <li key={i}><span style={{backgroundColor:CHART_COLORS[i]}}/>{series.name || `Series ${i + 1}`}</li>)}</ul>
        <div className="sheet-chart-controls"><label>Type<select aria-label={`Chart ${index + 1} type`} value={chart.chart_type} disabled={disabled} onChange={event => void onExecute([{kind:'chart.update',identity:chart.identity,expected_fingerprint_sha256:chart.fingerprint_sha256,chart_type:event.target.value as typeof type,title:chart.title,range:chart.range,anchor:chart.anchor}], 'Chart type updated')}>{options}</select></label><button disabled={disabled || !!reason} title={reason} onClick={() => void onExecute([{kind:'chart.update',identity:chart.identity,expected_fingerprint_sha256:chart.fingerprint_sha256,chart_type:chart.chart_type as typeof type,title:chart.title,range,anchor:chart.anchor}], 'Chart source updated')}>Use selected range</button><button disabled={disabled} onClick={() => void onExecute([{kind:'chart.delete',identity:chart.identity,expected_fingerprint_sha256:chart.fingerprint_sha256}], 'Chart deleted')}>Delete chart</button></div>
        <details className="sheet-chart-data"><summary>Chart data · {chart.categories.length} categories</summary><div><table><thead><tr><th>Category</th>{chart.series.map((series,i)=><th key={i}>{series.name}</th>)}</tr></thead><tbody>{chart.categories.map((label,i)=><tr key={i}><th>{label}</th>{chart.series.map((series,j)=><td key={j}>{series.values[i]}</td>)}</tr>)}</tbody></table></div></details>
      </> : <p className="sheet-chart-refusal">Preview and editing unavailable: {chart.refusal}. The chart is preserved in the workbook. Restore supported literal source values to re-enable charts created here.</p>}
    </section>)}
  </aside>;
}
