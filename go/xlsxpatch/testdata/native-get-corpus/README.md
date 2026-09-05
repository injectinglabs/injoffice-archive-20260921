# Native XLSX GET corpus

These are the same `.xlsx` bytes the InjOffice playground demo and the extract tests use.

They cover the Injecting native GET chain that previously 422/503'd in production:

| File | Expected native GET | Why it exists |
| --- | --- | --- |
| `pass-agent-dejavu.xlsx` | 200 | Agent-authored DejaVu table, no Excel view/xf noise |
| `pass-empty-inline-str.xlsx` | 200 | Blank `t="inlineStr"` cell with no `<is>` (Excel empty cell) |
| `pass-excel-defaults.xlsx` | 200 | Excel `showGridLines=1` and xf `pivotButton/quotePrefix=0` |
| `pass-excel-defaults-chart.xlsx` | 200 | The combined GET path: Excel defaults + DrawingML chart |
| `refuse-calibri.xlsx` | 422 `xlsx.native.font-unavailable` | Calibri is not a host-native face |
| `refuse-apply-flags.xlsx` | 422 `xlsx.native.unsupported` | cellXf ids differ without apply flags |
| `refuse-gridlines-off.xlsx` | 503 compile (`SHEET_VIEW_GEOMETRY`) | Explicit gridlines off is unbuilt view geometry |
| `refuse-freeze-pane.xlsx` | 503 compile (`SHEET_VIEW_GEOMETRY`) | Freeze pane is unbuilt view geometry |
| `refuse-quote-prefix.xlsx` | 503 compile (`STYLE_RECORD_ATTRIBUTES`) | `quotePrefix=1` is style-record authority |
| `refuse-shape-drawing.xlsx` | 422 `xlsx.native.unsupported` | Shapes are not chart-covered GET overlays |

Regenerate (do not hand-edit the packages):

```bash
go run ./go/xlsxpatch/cmd/nativegetcorpus
```

PNG/screenshots are not semantic pass/fail. The tests assert extract codes, drawing coverage, and (on the claw GET layer) HTTP status/refusal.
