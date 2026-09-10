# Charts, pivots, and formulas

The data-model helpers can be useful without embedding an editor. Optional manager, panel, and command integrations add UI behavior; your host still owns source data and persistence.

## Build chart data and aggregate a pivot

<<< @/examples/spreadsheet-tools.ts

The chart function produces an ECharts option, not a mounted chart or saved XLSX drawing. The pivot result is a deterministic grid, not an exported native pivot cache. Use the corresponding native conversion/write APIs when your workflow must persist these objects into a file.

The pivot input's first row supplies field names. The example groups West into a total of 35 and East into 35. Advanced Excel features such as slicers, arbitrary external sources, and general pivot semantic parity are outside this bounded example.

## Connect to an editor

The browser entries provide optional command registration and React panels. Supply a controller so edits use the same undo path. Enable a command only when you provide its required host—for example, an image-export host for chart exports or authenticated data-source callbacks for connectors.

These integrations do not silently provide persistence, offline synchronization, or authorization. Review [charts](../reference/generated/packages/charts), [pivots](../reference/generated/packages/pivots), and [Univer composition](../reference/generated/packages/univer-sheets).

## Formulas and calculation

`@injoffice/formulas` provides an audited function vocabulary and a facade over an explicitly selected engine. Writing formula text into native XLSX does not calculate it.

Client integrations register supported functions through the host editor's public facade. Server-friendly calculation jobs take source revisions and fingerprints; validate that a returned result still matches the current workbook before applying it.

The curated vocabulary is a regression scope, not proof of complete Excel formula semantics. [Formula integration reference](../reference/generated/packages/formulas) and [function inventory](../reference/generated/contracts/functions) describe the tested surface.

## Other spreadsheet features

- [Shapes](../reference/generated/packages/shapes): geometry, native conversion, and optional drawing controls.
- [Outlines](../reference/generated/packages/outlines): validated row/column grouping.
- [Sparklines](../reference/generated/packages/sparklines): small chart models and lifecycle commands.
- [Print](../reference/generated/packages/print): print configuration and host adapters.
- [Connectors](../reference/generated/packages/connectors): host-owned data sources, refresh jobs, and deterministic preprocessing.
