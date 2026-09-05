# Native XLSX playground fixture

`launch-readiness-plan.xlsx` is a deterministic, repository-owned fictional launch-readiness workbook. It contains six workstreams, named owners, statuses, budget and spend values, milestones, a styled summary row, and a native column chart. No external data or artwork is included.

The package, native workbook projection, and chart sidecar are generated together by `go/xlsxpatch/cmd/nativeplaygroundfixture`. Regenerate them from the repository root with:

```bash
go run ./go/xlsxpatch/cmd/nativeplaygroundfixture
```
