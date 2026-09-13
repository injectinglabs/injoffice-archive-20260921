# Authored preset geometry through a browser worker

`@injoffice/pptx-wasm` exposes the existing Go DrawingML catalog evaluator without requiring a PPTX package:

```ts
import { createPptxWasmClient } from '@injoffice/pptx-wasm'

const client = createPptxWasmClient()
try {
  const geometry = await client.evaluatePreset({
    name: 'wedgeRoundRectCallout',
    widthEmu: 4_000_000,
    heightEmu: 3_000_000,
    adjustments: { adj1: -50_000, adj2: 70_000 },
  })
  // Use geometry on an authored NativeShapeElement with this same frame.
  // Omit preset, retain paragraphs: [], and mark compatibility preserveOnly
  // with a warning diagnostic. The containing slide/deck propagate that status.
} finally {
  client.terminate()
}
```

The returned `NativeEvaluatedGeometry` works with the public native slide compiler and recording paint surface. Coordinates already include the requested frame dimensions. Resizing requires evaluating again; source guides and adjustments are not interpreted in JavaScript. Evaluation grants no save or source mutation authority. Existing mutation guards continue to refuse evaluated shapes.

The 187 names, guide semantics, arithmetic qualification, relative fill policy and explicit refusals are the same as [the source catalog](PPTX-PRESET-CATALOG.md). This API adds a caller-authored entry point, not broader geometry qualification or Office visual parity. Unsupported names, undeclared adjustments, ill-conditioned results and collapsed geometry refuse without poisoning the worker; a subsequent valid evaluation remains possible.

Requests accept exact case-sensitive fields `name`, `widthEmu`, `heightEmu` and optional `adjustments`. Frame dimensions are positive safe integer EMU; adjustment values are safe integers with negative zero refused. Client input is snapshotted before queueing, without reading accessors. Request JSON is limited to 64 KiB, adjustment objects to 1,024 entries, and response JSON to 2 MiB. The Go bridge rejects duplicate decoded field names, unknown/case-variant fields, nulls, noninteger JSON numbers and trailing data. Responses carry a protocol tag and request echo; the client binds that echo and validates all geometry through the existing public native contract before returning it.

The read-only operation shares initialization, operation IDs, queueing, cancellation, timeout and termination handling with other native WASM clients. Older engines can still initialize and extract/apply; asking an older engine to evaluate produces a recoverable refusal. Invalid response envelopes or geometry terminate the client as other malformed native responses do.

Validation includes all 187 presets through the Go JSON bridge, direct malformed request regressions, UTF-8 budgets, response mismatch/path validation, refusal recovery, and the XLSX/DOCX/PPTX runtime lifecycle suites. External functional evidence uses the public client in an actual classic Worker, the Go/WASM evaluator, and the existing `NativePptxVector` demo component. Arbitrary source rotation and the legacy file-preview integration remain separate geometry completion work.
