import type { Surface } from './route'

export type DemoSurface = Exclude<Surface, 'overview'>

export type DemoRecipeStep = {
  id: string
  title: string
  instruction: string
  evidence: string
}

export type DemoRecipeSource = {
  label: string
  path: string
  href: string
}

export type DemoRecipeRelation = {
  surface: DemoSurface
  reason: string
}

export type DemoRecipe = {
  id: string
  title: string
  outcome: string
  minutes: number
  steps: readonly DemoRecipeStep[]
  sources: readonly DemoRecipeSource[]
  related: readonly DemoRecipeRelation[]
}

const GITHUB_SOURCE_ROOT = 'https://github.com/injectinglabs/injoffice/blob/main/'

function source(label: string, path: string): DemoRecipeSource {
  return { label, path, href: `${GITHUB_SOURCE_ROOT}${path}` }
}

export const DEMO_RECIPES = {
  sheets: {
    id: 'sheets-native-round-trip',
    title: 'Prove an XLSX edit survives the round trip',
    outcome: 'Change a real workbook, reopen the emitted bytes, and inspect exact cell and revision readback evidence.',
    minutes: 3,
    steps: [
      { id: 'open-native', title: 'Open the native proof', instruction: 'Choose “Test XLSX round trip,” then select “Use bundled .xlsx.”', evidence: 'The workbook grid and a revision-bound Safe cell target appear.' },
      { id: 'edit-cell', title: 'Make one guarded edit', instruction: 'Choose a Safe cell, change its New literal value, and select “Save to XLSX.”', evidence: 'The status reports a successful mutation and verified download bytes become available.' },
      { id: 'verify-xlsx', title: 'Read the result back', instruction: 'In 03 · Verify, compare Cell, Before, After, and CAS moved before downloading.', evidence: 'After matches the value you entered and CAS moved shows the package revision changed.' },
    ],
    sources: [source('Native XLSX proof', 'apps/playground/src/pages/NativeRoundTripPage.tsx'), source('XLSX browser engine', 'packages/xlsx-wasm/src/index.ts')],
    related: [
      { surface: 'formulas', reason: 'Inspect the formula coverage behind calculated workbook cells.' },
      { surface: 'charts', reason: 'Turn a worksheet range into a renderer-neutral chart.' },
    ],
  },
  docs: {
    id: 'docs-guarded-text-edit',
    title: 'Perform a surgical DOCX text edit',
    outcome: 'Replace one exact text run without silently rebuilding the rest of the document.',
    minutes: 3,
    steps: [
      { id: 'load-docx', title: 'Load known DOCX bytes', instruction: 'Keep “In browser (default)” selected, then choose “Use bundled .docx.”', evidence: 'The preview renders document blocks while Fidelity boundary enumerates parsed content and explicit limitations.' },
      { id: 'choose-run', title: 'Choose a guarded run', instruction: 'Pick a Safe text run, change its Replacement text, and select “Save to DOCX.”', evidence: 'The save action is enabled only for a changed, bounded target.' },
      { id: 'verify-docx', title: 'Save and verify', instruction: 'In 03 · Verify, compare Before, After, identities, preserved parts, and CAS moved.', evidence: 'The target contains the new text and “Download verified .docx” becomes available.' },
    ],
    sources: [source('DOCX demo', 'apps/playground/src/pages/DocsPage.tsx'), source('DOCX browser engine', 'packages/docx-wasm/src/index.ts')],
    related: [
      { surface: 'history', reason: 'Compare document revisions as readable structured text changes.' },
      { surface: 'font-metrics', reason: 'Inspect the explicit text-layout boundary used by document hosts.' },
    ],
  },
  slides: {
    id: 'slides-outline-to-deck',
    title: 'Build and inspect an editable deck',
    outcome: 'Turn a plain outline into a themed DeckSpec and check its layout before native compilation.',
    minutes: 3,
    steps: [
      { id: 'revise-outline', title: 'Revise the outline', instruction: 'Edit the title and bullets in the Deck outline field.', evidence: 'The outline remains plain text and is ready to compile deterministically.' },
      { id: 'build-deck', title: 'Build the slides', instruction: 'Choose a Theme, select “Build slides,” then move through the deck with Prev and Next.', evidence: 'The canvas shows editable slides with stable content derived from the outline.' },
      { id: 'check-transition', title: 'Add motion and check layout', instruction: 'Set a transition for one slide, then review the Layout QC panel.', evidence: 'The transition is stored on the slide and QC reports any estimated overlap or overflow.' },
    ],
    sources: [source('Presentation demo', 'apps/playground/src/pages/SlidesPage.tsx'), source('Deck model', 'packages/slides/src/index.ts')],
    related: [
      { surface: 'pptx-authored', reason: 'Compile the DeckSpec into strict native presentation objects.' },
      { surface: 'pptx-render', reason: 'Inspect the renderer-neutral command stream for native slides.' },
    ],
  },
  pdf: {
    id: 'pdf-browser-edit',
    title: 'Edit a PDF without hiding the runtime boundary',
    outcome: 'Open real PDF bytes, apply a bounded browser-safe operation, and download the result.',
    minutes: 4,
    steps: [
      { id: 'open-pdf', title: 'Open a PDF', instruction: 'Wait for the bundled sample to load, or select “Open PDF” to use a local file, then page through it.', evidence: 'The canvas reports the current page and the PDF tool tabs become available.' },
      { id: 'inspect-pdf', title: 'Inspect its structure', instruction: 'Use Inspect to search text and read the outline, then open Markup or Forms for their document data.', evidence: 'Search, outline, and annotation results link to pages where available; form values stay explicit.' },
      { id: 'edit-pdf', title: 'Apply one page operation', instruction: 'Open Pages, then select “Rotate 90°” or “Insert blank after.”', evidence: 'The status describes the completed operation and the geometry or page count updates.' },
      { id: 'download-pdf', title: 'Export the edited bytes', instruction: 'Select “Download edited PDF” and retain the output as the proof artifact.', evidence: 'The downloaded file reflects the browser-local operation.' },
    ],
    sources: [source('PDF workbench', 'apps/playground/src/pages/PdfPage.tsx'), source('PDF operations', 'packages/pdf/src/index.ts')],
    related: [
      { surface: 'collab', reason: 'Try synchronized PDF annotations and page presence in two editors.' },
      { surface: 'history', reason: 'See how durable versions complement bounded document edits.' },
    ],
  },
  charts: {
    id: 'charts-range-to-render',
    title: 'Trace a cell range into a rendered chart',
    outcome: 'Edit source values and inspect both the live ECharts result and native chart wire.',
    minutes: 2,
    steps: [
      { id: 'choose-chart', title: 'Choose a chart type', instruction: 'Select a different chart type from the control strip.', evidence: 'The live preview redraws through the ECharts adapter.' },
      { id: 'edit-series', title: 'Change the source range', instruction: 'Edit one revenue value in the series table.', evidence: 'The preview and analysis values update from the same range.' },
      { id: 'inspect-wire', title: 'Inspect portable outputs', instruction: 'Compare Generated option with the Native XLSX wire panel.', evidence: 'Both outputs describe the selected type and current source range.' },
    ],
    sources: [source('Chart demo', 'apps/playground/src/pages/ChartsPage.tsx'), source('Chart adapter', 'packages/charts/src/index.ts')],
    related: [
      { surface: 'sheets', reason: 'Edit the workbook cells that commonly supply chart ranges.' },
      { surface: 'shapes', reason: 'Compare chart rendering with native drawing geometry.' },
    ],
  },
  pivots: {
    id: 'pivots-deterministic-summary',
    title: 'Build a deterministic pivot summary',
    outcome: 'Regroup a fixed record set and observe a reproducible, spreadsheet-ready result.',
    minutes: 2,
    steps: [
      { id: 'group-pivot', title: 'Change the row grouping', instruction: 'Switch Group rows between Region, Owner, and Quarter.', evidence: 'The generated row labels and totals update immediately.' },
      { id: 'aggregate-pivot', title: 'Change the measure', instruction: 'Choose Units or Revenue, then compare sum, average, and count.', evidence: 'The output grid applies the requested aggregation to the same source records.' },
      { id: 'filter-pivot', title: 'Filter and split', instruction: 'Choose one region and enable Split by quarter.', evidence: 'The pivot dimensions and generated columns reflect both controls.' },
    ],
    sources: [source('Pivot demo', 'apps/playground/src/pages/PivotsPage.tsx'), source('Pivot engine', 'packages/pivots/src/engine.ts')],
    related: [
      { surface: 'sheets', reason: 'Continue from a derived grid into a full spreadsheet workflow.' },
      { surface: 'charts', reason: 'Visualize the aggregated result through the chart adapter.' },
    ],
  },
  shapes: {
    id: 'shapes-fidelity-inspection',
    title: 'Inspect native shape fidelity',
    outcome: 'Find a native preset, customize its paint, and identify its preview fidelity class.',
    minutes: 2,
    steps: [
      { id: 'find-shape', title: 'Find a preset', instruction: 'Filter by category or search for a shape such as decision, star, or callout.', evidence: 'The preset library narrows while retaining native identifiers.' },
      { id: 'select-shape', title: 'Inspect the geometry', instruction: 'Select a result and compare its label, identifier, and large preview.', evidence: 'The inspector identifies whether the preview is distinct, approximate, or generic.' },
      { id: 'paint-shape', title: 'Change its paint', instruction: 'Choose new fill and stroke colors.', evidence: 'The same geometry redraws without changing its native preset identity.' },
    ],
    sources: [source('Shape demo', 'apps/playground/src/pages/ShapesPage.tsx'), source('Shape renderer', 'packages/shapes/src/index.ts')],
    related: [
      { surface: 'pptx-native', reason: 'Apply an exact AutoShape edit inside real PPTX bytes.' },
      { surface: 'pptx-render', reason: 'Follow native geometry into ordered paint commands.' },
    ],
  },
  connectors: {
    id: 'connectors-payload-to-range',
    title: 'Normalize host data into a bound range',
    outcome: 'Transform JSON or CSV into a deterministic grid without moving credentials into the document.',
    minutes: 2,
    steps: [
      { id: 'choose-payload', title: 'Choose a source format', instruction: 'Switch between JSON and CSV in the control strip.', evidence: 'A valid host-supplied sample and matching parser controls appear.' },
      { id: 'edit-payload', title: 'Edit the payload', instruction: 'Change a value or add a row; for JSON, try a different dot path.', evidence: 'The spreadsheet-ready grid updates as you type or reports a precise parse issue.' },
      { id: 'inspect-pipeline', title: 'Trace preprocessing', instruction: 'Review the Range preprocess stages beside the normalized grid.', evidence: 'The stage list explains the deterministic path from payload to bound values.' },
    ],
    sources: [source('Connector demo', 'apps/playground/src/pages/ConnectorsPage.tsx'), source('Connector pipeline', 'packages/connectors/src/index.ts')],
    related: [
      { surface: 'sheets', reason: 'Place normalized connector results into a live workbook.' },
      { surface: 'pivots', reason: 'Aggregate the resulting tabular records into a pivot.' },
    ],
  },
  formulas: {
    id: 'formulas-audit-to-calculation',
    title: 'Move from audited coverage to a calculation job',
    outcome: 'Find a supported business function, inspect its fixture, and run it through the calculator.',
    minutes: 2,
    steps: [
      { id: 'find-formula', title: 'Find an audited function', instruction: 'Search for a function family such as lookup, date, or text.', evidence: 'The inventory narrows while preserving each function’s audited status.' },
      { id: 'inspect-fixture', title: 'Inspect a valid invocation', instruction: 'Select one result and read its fixture formula and compatibility details.', evidence: 'The detail panel shows a concrete expression rather than a name-only claim.' },
      { id: 'calculate-formula', title: 'Run the job', instruction: 'Select “Use selected formula,” then select “Submit job.”', evidence: 'The page returns a value or an explicit engine error with no silent fallback.' },
    ],
    sources: [source('Formula demo', 'apps/playground/src/pages/FormulasPage.tsx'), source('Formula calculation API', 'packages/formulas/src/calculation.ts')],
    related: [
      { surface: 'sheets', reason: 'Use audited functions in the live spreadsheet editor.' },
      { surface: 'history', reason: 'Inspect readable diffs when formulas or cell values change.' },
    ],
  },
  agent: {
    id: 'agent-guarded-change-set',
    title: 'Run a guarded agent change from proposal to proof',
    outcome: 'Use Sheets for a real XLSX write, reopen, and download; explore the same guarded lifecycle with simulated Docs, Slides, and PDF data.',
    minutes: 3,
    steps: [
      { id: 'prepare-agent-change', title: 'Run the AI-first proposal', instruction: 'Start with Sheets for a real file. Try “Mark Mobile as On track” and choose “Run agent.” Local mode needs no model. Optional live mode requires a configured host and consent. Other formats remain lifecycle simulations.', evidence: 'The actual tool trace includes capability discovery and bounded workbook reads. Preview never replaces the source file.' },
      { id: 'review-agent-diff', title: 'Review the exact change', instruction: 'Compare the artifact preview, proposed change, semantic diff, source identity, and advertised operations.', evidence: 'The proposal stays bound to one source revision and only uses a discovered capability.' },
      { id: 'approve-agent-commit', title: 'Approve, verify, and download', instruction: 'Approve the exact diff, then commit and download. Try the safety controls on a fresh sample: another editor invalidates the plan; retry returns the same receipt; injected readback failure withholds a verified download.', evidence: 'The public adapter verifies the requested target in a fresh native snapshot. Post-write verification is part of the commit receipt, not a second office.verify call.' },
    ],
    sources: [source('Agent workflow demo', 'apps/playground/src/pages/AgentPage.tsx'), source('Real-file XLSX adapter', 'apps/playground/src/agentXlsxDemo.ts'), source('Agent change-set core', 'packages/agent-tools/src/session.ts')],
    related: [
      { surface: 'collab', reason: 'Continue into real-time human and agent participation over shared operations.' },
      { surface: 'history', reason: 'Capture verified agent output as an attributable durable version.' },
    ],
  },
  collab: {
    id: 'collab-two-editor-proof',
    title: 'Prove an operation crosses editor boundaries',
    outcome: 'Make an edit on one surface, observe it on another, and inspect the ordered operation ledger.',
    minutes: 3,
    steps: [
      { id: 'choose-collab-format', title: 'Choose an artifact', instruction: 'Keep “Two-editor simulation” selected, then choose Sheets, Docs, Slides, or PDF.', evidence: 'Two independent editors join one browser-local room.' },
      { id: 'make-collab-edit', title: 'Edit on one side', instruction: 'Change content or presence in the left editor.', evidence: 'The right editor receives the operation while retaining its own local identity.' },
      { id: 'inspect-ledger', title: 'Inspect ordering', instruction: 'Read the Shared operation ledger and compare ordered content operations with presence events.', evidence: 'Ordered edits have sequence numbers while ephemeral presence is labeled separately.' },
    ],
    sources: [source('Two-editor collaboration proof', 'apps/playground/src/collabSimulator.tsx'), source('Collaboration core', 'packages/collab/src/index.ts')],
    related: [
      { surface: 'history', reason: 'Persist meaningful collaboration checkpoints as durable versions.' },
      { surface: 'pdf', reason: 'Explore the full browser PDF tools behind the shared annotation proof.' },
    ],
  },
  history: {
    id: 'history-diff-and-restore',
    title: 'Capture, explain, and restore a version',
    outcome: 'Turn two snapshots into a readable diff, capture attribution, and restore without rewriting history.',
    minutes: 3,
    steps: [
      { id: 'edit-snapshots', title: 'Create a meaningful change', instruction: 'Choose Spreadsheet or Document and edit the Before and After content.', evidence: 'The Changes panel reports structured cells or text spans rather than opaque bytes.' },
      { id: 'capture-version', title: 'Capture with attribution', instruction: 'Select Capture after or Capture as agent.', evidence: 'A new timeline entry records the author and capture reason.' },
      { id: 'restore-version', title: 'Restore non-destructively', instruction: 'Select a version and choose Restore selected.', evidence: 'The restored snapshot becomes a new version while earlier history stays intact.' },
    ],
    sources: [source('History demo', 'apps/playground/src/pages/HistoryPage.tsx'), source('History manager', 'packages/history/src/manager.ts')],
    related: [
      { surface: 'docs', reason: 'Apply an exact document edit before explaining its revision.' },
      { surface: 'collab', reason: 'See the live operation stream that can feed durable captures.' },
    ],
  },
  'font-metrics': {
    id: 'font-metrics-layout-boundary',
    title: 'Test a conservative line-break decision',
    outcome: 'Change adjacent text clusters and see the browser-safe contract separate from native shaping.',
    minutes: 2,
    steps: [
      { id: 'edit-clusters', title: 'Change the cluster boundary', instruction: 'Edit the left and right clusters around the visible boundary marker.', evidence: 'The break decision updates for the exact pair of Unicode clusters.' },
      { id: 'scale-type', title: 'Scale the run', instruction: 'Move the font-size control between 8 and 72 points.', evidence: 'Integer milli-point inputs and the scaled ascent update together.' },
      { id: 'inspect-boundary', title: 'Read the runtime boundary', instruction: 'Compare the browser-safe decision with the Host runtime boundary note.', evidence: 'The page names the Node-only resolver, font-byte, shaping, and bidi responsibilities.' },
    ],
    sources: [source('Typography demo', 'apps/playground/src/pages/FontMetricsPage.tsx'), source('Font metrics contract', 'packages/font-metrics/src/index.ts')],
    related: [
      { surface: 'docs', reason: 'See where conservative text layout feeds document rendering.' },
      { surface: 'slides', reason: 'Review layout quality checks for authored presentation text.' },
    ],
  },
  'pptx-authored': {
    id: 'pptx-authored-native-contract',
    title: 'Compile DeckSpec into the native contract',
    outcome: 'Change authored content and inspect the immutable native presentation objects it produces.',
    minutes: 2,
    steps: [
      { id: 'edit-authored', title: 'Edit authored inputs', instruction: 'Change the deck title, subtitle, or authored theme.', evidence: 'The compiler runs immediately from plain DeckSpec data.' },
      { id: 'verify-contract', title: 'Verify the contract', instruction: 'Check the acceptance state, contract version, origin, and native element count.', evidence: 'Valid input produces pptx-native/v1 data; invalid input is explicitly refused.' },
      { id: 'inspect-slide', title: 'Inspect one slide', instruction: 'Choose another compiled slide from the inspector.', evidence: 'The preview lists the exact native element kinds for that stable slide identity.' },
    ],
    sources: [source('PPTX authoring demo', 'apps/playground/src/pages/PptxAuthoredPage.tsx'), source('Deck compiler', 'packages/pptx-authored/src/compiler.ts')],
    related: [
      { surface: 'slides', reason: 'Build and edit the higher-level DeckSpec used as compiler input.' },
      { surface: 'pptx-render', reason: 'Continue from native objects into renderer-neutral commands.' },
    ],
  },
  'pptx-native': {
    id: 'pptx-native-guarded-edit',
    title: 'Round-trip one exact PPTX mutation',
    outcome: 'Edit a guarded text run or AutoShape and verify the reopened presentation bytes.',
    minutes: 3,
    steps: [
      { id: 'load-pptx', title: 'Load native PPTX bytes', instruction: 'Use the bundled .pptx with the default in-browser runtime.', evidence: 'The exact native projection lists slides and supported mutation targets.' },
      { id: 'choose-pptx-target', title: 'Choose one exact target', instruction: 'Select a text or AutoShape target, then change only the offered fields.', evidence: 'The mutation remains bounded to a revision-bound element and operation kind.' },
      { id: 'verify-pptx', title: 'Save and reopen', instruction: 'Save to PPTX and inspect the current mutation evidence.', evidence: 'The reopened output matches the requested value and remains downloadable as PPTX bytes.' },
    ],
    sources: [source('Native PPTX demo', 'apps/playground/src/pages/PptxNativePage.tsx'), source('PPTX browser engine', 'packages/pptx-wasm/src/index.ts')],
    related: [
      { surface: 'pptx-render', reason: 'Render native slide objects into a deterministic command stream.' },
      { surface: 'shapes', reason: 'Browse the broader catalog of supported native shape identifiers.' },
    ],
  },
  'pptx-render': {
    id: 'pptx-render-command-stream',
    title: 'Record deterministic slide paint commands',
    outcome: 'Change native geometry and inspect an ordered command stream independent of any renderer.',
    minutes: 2,
    steps: [
      { id: 'choose-preset', title: 'Choose native geometry', instruction: 'Select a different Native shape preset.', evidence: 'The host SVG preview redraws from the selected preset path.' },
      { id: 'toggle-connector', title: 'Add a connector', instruction: 'Toggle “Include a connector in the command stream.”', evidence: 'The command count and ordered list update deterministically.' },
      { id: 'inspect-commands', title: 'Inspect paint order', instruction: 'Read the command kinds and their source element identities from top to bottom.', evidence: 'The recording separates native compilation from the host’s eventual paint implementation.' },
    ],
    sources: [source('PPTX render demo', 'apps/playground/src/pages/PptxRenderPage.tsx'), source('Render command layer', 'packages/pptx-render/src/index.ts')],
    related: [
      { surface: 'pptx-authored', reason: 'Generate the native slide objects consumed by render compilation.' },
      { surface: 'pptx-native', reason: 'Extract and mutate native objects from real PPTX bytes.' },
    ],
  },
} as const satisfies Record<DemoSurface, DemoRecipe>
