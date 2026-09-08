# @injoffice/agent-tools

Model-neutral agent sessions and revision-safe change sets for InjOffice artifacts. The package contains no model client, server, or UI. Applications supply a format adapter and keep control of authentication, authorization, persistence, and human approval.

The same workflow applies to every adapter-backed format:

```text
capabilities -> inspect/read -> plan -> validate/preview/diff -> commit -> verify
```

## Session API

```ts
import { createAgentSession } from '@injoffice/agent-tools'

// Private host state, populated only by the authenticated review action.
// Never expose a tool that lets the model add entries to this set.
const approvedPlans = new Set<string>()

const session = await createAgentSession({
  artifact,
  adapter,
  actor: { id: 'forecast-agent', kind: 'agent', displayName: 'Forecast assistant' },
  confirmDestructive: async ({ action, changeSet }) =>
    action === 'commit' && !!changeSet && approvedPlans.has(JSON.stringify(changeSet)),
})

const { capabilities } = await session.capabilities()
const context = await session.inspect({ query: { scope: 'sheet:Revenue' }, maxItems: 100 })
const changeSet = await session.plan([
  // The selected adapter defines the operation name and input schema.
  { name: 'your-format.operation', input: { targetId: 'target-1', value: 42 } },
])

await changeSet.validate()
await changeSet.preview()
await changeSet.diff()
// In the trusted host's authenticated approval handler, after presenting this
// exact plan to the authorized reviewer (not in response to model arguments):
approvedPlans.add(JSON.stringify(changeSet.envelope))
const committed = await changeSet.commit({ idempotencyKey: 'forecast-run-42' })
// committed.verification was produced against the replacement artifact.
```

`plan` returns an immutable, plain-JSON envelope bound to the artifact revision and fingerprint. Commits are serialized, idempotent within the session, rejected when stale, and passed to the adapter with a compare-and-swap contract. Destructive capabilities require the host confirmation hook. The session adopts a replacement artifact only after its adapter-reported identity has been independently read back.

The in-memory approval set above illustrates the trust boundary, not a production authorization store. Bind approvals to the exact plan, authenticated principal, artifact identity, and source revision; scope their lifetime to that review/session. Perform authorization at the host boundary for **all** writes, including operations that do not advertise confirmation. Do not authorize a mutation merely because the model supplied `confirmation: "approved"`. Hosts must also authenticate the actor and provide persistent atomic revision checks and durable idempotency where artifacts outlive the process.

Adapters advertise every operation with a JSON-Schema-compatible `inputSchema`. They implement format-specific inspection, mutation, rendering, diffing, and verification while the package enforces common limits and lifecycle rules.

## Transport-neutral tools

`createAgentToolDispatcher(session)` exposes JSON-only descriptors and accepts `AgentToolCall` values. It retains a bounded set of planned change sets by opaque `changeSetId`, so an MCP server, HTTP route, worker bridge, or any model framework can use the protocol without importing model-specific code.

```ts
import {
  AGENT_TOOLS_PROTOCOL,
  AGENT_TOOLS_PROTOCOL_VERSION,
  createAgentToolDispatcher,
} from '@injoffice/agent-tools'

const tools = createAgentToolDispatcher(session, { maxChangeSets: 32 })
const result = await tools.dispatch({
  protocol: AGENT_TOOLS_PROTOCOL,
  protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
  requestId: 'request-1',
  method: 'office.capabilities',
  params: {},
})
```

The built-in tool surface is `office.capabilities`, `office.inspect`, `office.read`, `office.plan`, `office.validate`, `office.preview`, `office.diff`, `office.commit`, `office.verify`, and optional `office.restore`. Call `forget(changeSetId)` to release a retained plan early.

`office.verify` verifies the **planned** result. `office.commit` independently verifies the committed replacement and returns that result in its receipt. Do not present a planned verification as proof of saved output. A completed write with failed verification is not the same as a rejected pre-write plan, and must not be presented as a verified download or an automatic rollback.

The [playground Sheets example](../../apps/playground/README.md#agent-integration-example) exercises the public dispatcher and XLSX adapter with real native browser I/O, an editable deterministic proposal, trusted local approval, and visible stale-revision/retry/verification-failure scenarios. Its local proposal is not a model integration.

## Adapter responsibilities

- Return plain JSON from every method except the adapter-private replacement `artifact`.
- Keep artifact IDs and format IDs stable across revisions.
- Enforce `expectedRevision` and `expectedFingerprint` atomically inside `commit` and `restore`.
- Durably deduplicate idempotency keys when artifacts are persisted outside this process.
- Return a replacement artifact and identity from mutations. Immutable byte arrays/specs are supported.
- Verify committed output from the exact replacement artifact passed in `AgentAdapterVerifyRequest`.
- Preserve or explicitly refuse unsupported native document features.

This package is browser-safe and dependency-free. Abort signals remain local control values and never enter a wire envelope.
