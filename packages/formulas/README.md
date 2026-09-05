# @injoffice/formulas

A typed compatibility inventory for spreadsheet functions that require explicit host coverage.

```bash
npm install @injoffice/formulas
```

```ts
import { TARGET_FUNCTIONS } from '@injoffice/formulas'

for (const fn of TARGET_FUNCTIONS) {
  console.log(fn.name, fn.formula)
}
```

## Client formula facade

`ClientFormulaIntegration` attaches the audited client vocabulary and
host-defined functions to Univer's public formula facade without replacing its
engine:

```ts
import {
  ClientFormulaIntegration,
  createAdvancedFormulaFeatureProvider,
} from '@injoffice/formulas'

const client = ClientFormulaIntegration.fromUniver(univerAPI, [{
  name: 'ACME.RATE',
  description: 'Returns an authorized host rate.',
  requiredCapabilities: ['formula.acme-rate'],
  calculate: (basis) => Number(basis) * 0.05,
}])

// Structurally compatible with the advancedFormula feature-registry slot.
const provider = createAdvancedFormulaFeatureProvider(client)
const loaded = await provider.load()
const activation = loaded.activate({
  serverCapabilities: new Set(['formula.acme-rate']),
})

activation.setFormula(worksheet.getRange('A1'), '=LET(x,2,x*3)')
await activation.recalculate()
activation.dispose() // unregisters ACME.RATE
```

Definitions are validated and sorted before registration. Duplicate names,
unknown allow-list entries, missing required capabilities, unavailable async
registration, and accidental overrides of audited built-ins fail before the
facade changes. A later registration failure rolls back earlier functions;
normal disposal unregisters in reverse order and is idempotent. Hosts may omit
capability-gated functions explicitly instead of failing the activation.

`targetVocabulary` exposes the same 142 names as `TARGET_FUNCTIONS`; the real
headless-engine test writes and recalculates every sample through this facade.
That is a pinned regression gate for the curated vocabulary, not proof of every
Excel function, worker/performance equivalence, volatile semantics, dependency
behavior, or error/debugging parity with Univer Pro.

The package also provides a server-friendly lifecycle around a formula engine selected by your host:

```ts
import { CalculationManager, assertCalculationResultCurrent } from '@injoffice/formulas'

const calculations = new CalculationManager({
  id: 'my-formula-engine',
  version: '2026-09',
  deterministic: true,
  async calculate(context) {
    // Evaluate context.snapshot/context.formulas in a worker or remote service.
    return { cells: [] }
  },
})

const job = calculations.submit({
  jobId: 'calc-42',
  workbookId: 'budget',
  sourceRevision: 'revision-17',
  sourceFingerprint: 'sha256:...',
  snapshot: workbookSnapshot,
  formulas: [{ sheetId: 'sheet-1', row: 0, column: 1, formula: '=SUM(A1:A10)' }],
  expectedEngineFingerprint: calculations.engineFingerprint,
})

const result = await job.result
assertCalculationResultCurrent(result, {
  workbookId: 'budget', revision: currentRevision, fingerprint: currentFingerprint,
})
```

## Collaborative results

`FormulaCollaborationSession` distributes validated derived results through a
host transport without treating formula-engine internals as ordinary workbook
mutations. A negotiated authority publishes a monotonic result stream; followers
verify the exact authority epoch and workbook revision before one atomic apply:

```ts
import { FormulaCollaborationSession } from '@injoffice/formulas'

const sharedResults = new FormulaCollaborationSession({
  room: 'budget',
  authorityId: 'calculation-primary',
  authorityEpoch: 'lease-42',
  role: 'follower',
  transport: resultTransport,
  target: {
    currentIdentity: () => workbookIdentity,
    inspect: (addresses) => worksheet.inspectFormulaOccupancy(addresses),
    applyDerived: (writes, envelope) => worksheet.applyDerivedAtomically(writes, envelope.sequence),
  },
  onRecalculationRequired: (reason) => requestAuthoritativeCalculation(reason),
  onResyncRequired: () => reloadCalculationSequence(),
})

sharedResults.start()
```

The spill planner requires every anchor to remain a formula, preserves only
cells already owned by that anchor, and refuses dirty, occupied, missing, or
overlapping destinations. Stale and conflicting envelopes advance the ordered
stream but trigger recalculation; gaps and apply failures block until an exact-
epoch resync. The host still owns authority election, authenticated transport,
durable sequence enforcement, formula-edit replication, revision assignment,
and atomic worksheet writes.

The protocol canonicalizes formula and metadata ordering, validates scalar/error/spill output, rejects stale workbook revisions, and provides cancellation, timeouts, lifecycle events, and engine-contract negotiation. The host must fingerprint the complete snapshot and remains responsible for transport, authorization, persistence, and applying derived results.

This package does not contain a formula evaluator. `TARGET_FUNCTIONS` is an auditable bounded compatibility matrix, and `CalculationManager` is orchestration around a host-injected engine. Neither is a claim of Excel or Univer Pro calculation-engine equivalence.
