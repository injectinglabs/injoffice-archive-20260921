# Agent change sets

InjOffice provides a provider-independent execution boundary for document agents. A model, deterministic rule, or human can propose plain operations; InjOffice inspects the target, creates an isolated change set, validates it, applies an approved commit against an expected revision, and verifies the output.

```text
inspect -> plan -> preview + diff -> validate -> approve -> commit -> verify
```

The core packages do not host a model, proxy model traffic, keep API keys, or require an InjOffice service.

## Architecture

`@injoffice/agent-tools` defines the common session, change-set, evidence, limits, refusal, registry, and adapter contracts. `@injoffice/agent-office` provides format adapters. Applications pass an adapter directly or register adapters in `AgentAdapterRegistry`; they also provide the model or other proposal source and the approval policy.

An adapter is responsible for:

1. Reporting exact operations supported for an artifact.
2. Returning bounded inspections tied to a revision and fingerprint.
3. Translating supported semantic operations into native format mutations.
4. Producing an isolated preview and structured diff without changing the source.
5. Refusing unsupported, ambiguous, stale, or policy-blocked work.
6. Applying an approved plan atomically where the underlying engine permits it.
7. Reopening the output and returning format-specific verification evidence.

## Safety contract

- Agents discover capabilities instead of assuming Office parity.
- Reads and operation batches have host-configured limits.
- A change set is immutable and records its source revision.
- Commit requires a matching expected revision and an idempotency key.
- Destructive operations can require an explicit host confirmation callback.
- Structured issues include stable codes and JSON paths.
- A refusal returns no partial output.
- Verification is separate from successful mutation and carries output evidence.

The host still owns identity, authorization, tenant isolation, durable storage, malware scanning, model policy, and the user approval experience.

## Formats

The packages do not contain a model SDK. The workflow is format-neutral, but capabilities are not:

- **XLSX:** bounded inspection of native v1/v2 projections plus v1 cell, formula, style, dimension, and merge mutations.
- **DOCX:** bounded document inspection and semantic guarded text replacement.
- **PPTX:** authored-deck slide operations plus source-anchored native text and AutoShape updates.
- **PDF:** bounded page inspection and rotate, insert, delete, reorder, crop, resize, or n-up operations.

Unsupported macros, encrypted packages, embedded objects, ambiguous targets, or out-of-scope formatting must be preserved untouched or refused.

New formats implement the same adapter contract. They do not require changes to an agent loop or provider integration.

## Non-goals

This layer is not:

- a built-in chat application or autonomous agent;
- a required hosted backend or model proxy;
- an abstraction that claims every document supports the same mutations;
- a replacement for application authorization or human-review policy;
- permission to silently reconstruct unsupported native files.

## Playground proof

The `#/agent` playground route demonstrates capability discovery, a bounded inspection, an operation proposal, isolated preview and diff, validation, explicit human approval, commit, and verification. The default built-in mock agent needs no model service, API key, or endpoint configuration. It uses a same-origin mock endpoint in development and the identical deterministic generator in-browser on static builds. Its proposal is explicitly simulated; its refusal mode proves that unsupported operations stop before write. Local rule-based and live proposals remain optional; only live proposals require a configured local host and explicit consent. See [proposal modes](AGENT-PROPOSAL-HOST.md).

The **Sheets** example loads the bundled `launch-readiness-plan.xlsx`. Editable requests such as “Mark Security as Ready” and “Mark Mobile as On track” discover the correct target through public capability, inspection, and read calls. The shipped XLSX adapter delegates preview and write to native browser Worker callbacks. Preview applies to an isolated copy, never the authoritative source. After host approval, commit writes the source replacement and verifies native readback. Verified output bytes can be downloaded. The displayed table is a bounded projection, not a rendered Excel page. Sample and engine assets are fetched from the demo host; Office file bytes are never uploaded and no native-write server fallback is attempted. Live proposals share only the displayed bounded context after consent.

**Docs** loads `northstar-launch-brief.docx`, discovers editable source-anchored text through public reads, and replaces one exact run using the native DOCX adapter and browser Worker. **Slides** loads the populated launch review as `northstar-quarterly-review.pptx` and replaces one exact native text element through the native PPTX adapter and Worker. Both preview copied file bytes, bind approval to the exact plan, write a source replacement, and re-extract the output before offering a verified download. Their text views are bounded projections, not full-fidelity page or slide rendering.

**PDF** creates the populated `northstar-operating-review.pdf` fixture locally and rotates one disclosed page through the public PDF adapter. Its preview shows parsed page metadata. Commit reparses the real output bytes and checks requested rotations, page count, geometry, and the output fingerprint. PDF execution uses browser PDF code, not a Go Worker. No text rewriting or arbitrary PDF edits are implied.

All four examples default to a clearly labelled deterministic mock proposal, not an LLM. File execution and downloaded bytes are real. Each demonstrates approval, stale-source rejection, idempotent retry, unsupported-operation refusal, and a labelled post-write verification failure that withholds the verified download.

### Minimal host integration

```ts
import { createAgentSession } from '@injoffice/agent-tools'

// artifact, adapter, actor and proposedOperations come from your host.
// The adapter owns format I/O; approval must come from your review UI.
const approvedChanges = new Set<string>() // private to the trusted host
const session = await createAgentSession({
  artifact, adapter, actor,
  confirmDestructive: async ({ changeSet }) => !!changeSet && approvedChanges.has(changeSet.changeSetId),
})
const change = await session.plan(proposedOperations, {
  expectedRevision: session.identity.revision,
  expectedFingerprint: session.identity.fingerprint,
})
const preview = await change.preview()
const diff = await change.diff()
const validation = await change.validate()
// Show preview/diff, then wait for explicit approval of this exact change.
if (validation.valid && approvedByUser) {
  approvedChanges.add(change.envelope.changeSetId) // trusted host review action only
  const receipt = await change.commit({ idempotencyKey })
  // A completed mutation and successful verification are distinct outcomes.
  showResult(receipt.verification)
}
```

Open the proofs for [#/agent?format=sheets](#/agent?format=sheets), [#/agent?format=docs](#/agent?format=docs), [#/agent?format=slides](#/agent?format=slides), or [#/agent?format=pdf](#/agent?format=pdf).
