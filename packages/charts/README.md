# @injoffice/charts

Renderer-neutral chart specifications, data extraction, ECharts options, and a Univer floating-chart integration.

```bash
npm install @injoffice/charts echarts react react-dom
```

```ts
import { buildEChartsOption, extractChartData, type ChartSpec } from '@injoffice/charts'

const values = [
  ['Month', 'Revenue'],
  ['Jan', 120],
  ['Feb', 165],
]
const spec: ChartSpec = {
  id: 'revenue',
  type: 'Column',
  title: 'Revenue',
  range: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
}

const option = buildEChartsOption(spec, extractChartData(values, spec))
```

The pure functions work without Univer or the DOM. `ChartManager`, `ChartFloat`, and `ChartPanel` provide the optional editor integration. A host owns workbook persistence and supplies the sheet values referenced by each `ChartSpec`.

`ChartCommandController` adds live handles and complete before/after snapshots
for create, add, update, remove, and per-sheet layering. Snapshot chart arrays
are ordered back-to-front. `bringForward`, `sendBackward`, `bringToFront`, and
`sendToBack` change only charts on the same worksheet. An optional
`ChartLayerHost` applies the accepted order to an already-mounted visual layer;
without one, the model order takes effect when the snapshot is next mounted.
Browser hosts can import
`@injoffice/charts/browser` and call `registerUniverChartCommands` to register
configurable public Univer commands plus native undo/redo restore mutations.
The optional Insert ribbon button is registered only when the host supplies a
chart-picker callback. Snapshot restore validates ids, types, ranges, anchors,
and native identities before replacing manager state, and attempts to restore
the prior complete snapshot if a chart cannot be mounted.

Pass the returned `controller` to `<ChartPanel controller={controller} />` so
panel edits, removal, and the four built-in layer controls enter that same undo
path. `showLayerControls` can omit the layer UI independently. Passing the
optional `exporter` returned by the browser registration enables a PNG, JPEG,
or SVG picker and export button; `showExportControls` hides it independently.
Validated artifacts go to `onExport`, or to a user-initiated browser download
when no callback is supplied. The legacy `manager` prop remains available for
hosts that deliberately omit an undo sink.

## Collaboration

`ChartCollaborationSession` is a transport-neutral protocol for server-ordered
chart-object changes. A host sends an intent to its collaboration authority,
assigns a monotonically increasing workbook revision, and broadcasts the
resulting event. Create, update, remove, and layer events target one stable
`ChartSpec.id`; they never contain a complete chart collection and remote
events do not enter the local snapshot-undo stack. Stable native identities
cannot be replaced by an update.

Each chart has an object revision. Each worksheet also has a layer revision,
because back-to-front order is relational: creating, removing, or layering one
chart invalidates concurrently-authored layer operations on that sheet. The
authority hook is host-injected and revisioned, so a permission-policy change
cannot be silently evaluated under stale rules. Denials and deterministic base
revision conflicts consume their assigned stream revision. Stream gaps,
authority-revision mismatches, hook failures, and visual-host application
failures stop application and request resynchronization.

`revisionState()` contains revision metadata and removal tombstones only. The
host remains responsible for durable transport and for hydrating the matching
chart collection through `ChartManager` before `resync()`. This bounded model
does not provide offline queues, operation-log storage, property-level merging,
cross-sheet chart moves, collaborative undo transformation, or an OT/CRDT
implementation.

## Image export

`ChartImageExportManager` supplies cancellable, timeout-bounded image export
jobs for `png`, `jpeg`, and `svg`. The caller injects a `ChartImageExportHost`;
InjOffice passes it an immutable chart snapshot and `AbortSignal`, validates the
returned media type, byte budget, and dimensions, clones the bytes, and emits
metadata-only `started`, `completed`, `canceled`, `timed-out`, or `failed`
events. Cancellation and timeout settle the public job even when the renderer
ignores its signal. The package does not claim that exports from different host
renderers are pixel-identical.

`registerUniverChartCommands` registers `layer` when a `layerHost` is supplied
(or when a host explicitly enables model-only layering). It registers
`exportImage` only when `imageExport.host` is supplied, so applications can
omit either capability from their command surface and bundles. The panel never
selects a renderer: it only exposes export when the host supplies one, and file
delivery occurs only after the user presses Export.

## Chart vocabulary

The renderer-neutral model covers all 30 chart families in the current public
Univer Charts guide: `Line`, `Column`, `ColumnStacked`,
`ColumnPercentStacked`, `Bar`, `BarStacked`, `BarPercentStacked`, `Pie`,
`Donut`, `Area`, `AreaStacked`, `AreaPercentStacked`, `Radar`, `Scatter`,
`Combination`, `WordCloud`, `Funnel`, `Bubble`, `Relation`, `Waterfall`,
`Pareto`, `Sankey`, `Heatmap`, `Boxplot`, `Candlestick`, `Histogram`,
`Treemap`, `Sunburst`, `Gauge`, and `Chord`.

Snapshots using the original lowercase names (`column`, `doughnut`,
`waterfall`, and the other values in `LEGACY_CHART_TYPES`) remain supported.
Use `normalizeChartType` at persistence boundaries and
`validateChartInput`/`assertValidChartInput` for untrusted agent or RPC input.

Most plots use categories plus one or more numeric series. Specialty layouts
have explicit conventions:

- `Bubble`: first three series are x, y, and size.
- `Candlestick`: first four series are open, close, low, and high.
- `Relation`, `Sankey`, and `Chord`: categories are `Source → Target`; the
  first series supplies edge weights.
- `Sunburst`: categories are slash-delimited paths; the first series supplies
  leaf values.
- `Histogram`: first series contains raw observations.
- `WordCloud`, `Gauge`, `Pie`, `Donut`, `Funnel`, `Treemap`, `Waterfall`, and
  `Pareto`: first series supplies values.
- `Boxplot`: every series is summarized independently.

`WordCloud` uses a deterministic ECharts custom series and `Chord` uses the
standard circular graph series, so neither requires an optional renderer
plugin. They preserve the public chart intent and data contract but are not a
pixel-identical claim against another renderer.

## Native XLSX boundary

ECharts rendering coverage is not a native-file fidelity claim. The native
writer currently supports only `Line`, `Column`, `Bar`, `Area`, `Pie`, `Donut`,
and `Scatter` (including their legacy names). `toWireCharts` explicitly skips
the other chart types with a diagnostic; it never silently substitutes a
different OOXML plot.

Hydrated native charts carry `NativeChartIdentity` (chart part, drawing part,
and DrawingML `cNvPr` object id). `toWireChartUpdate` and
`toWireChartRemove` build identity-bound requests; `toWireCharts` refuses to
serialize a hydrated chart as a duplicate add. The Go `xlsxpatch.UpdateChart`
and `RemoveChart` operations retain sibling anchors and relationships and
fail closed for stale or ambiguous identities, non-two-cell targets, shared
chart parts, or chart dependency graphs the bounded writer cannot preserve.
Layer order and exported image bytes are not written back to native XLSX by
this slice; native DrawingML z-order remains unchanged unless a future bounded
native reorder operation explicitly supports it.
