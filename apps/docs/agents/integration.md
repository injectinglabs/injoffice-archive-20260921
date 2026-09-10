# Tool protocol and adapters

`@injoffice/agent-tools` provides a model-neutral session and JSON dispatcher. `@injoffice/agent-office` translates the common lifecycle into format-specific capabilities. Your application owns the transport and execution environment.

## Tool lifecycle

| Tool | Purpose | Writes the source? |
| --- | --- | --- |
| `office.capabilities` | Discover supported operations and schemas | No |
| `office.inspect`, `office.read` | Retrieve bounded source context | No |
| `office.plan` | Retain an immutable revision-bound change set | No |
| `office.validate` | Check the proposed operation set | No |
| `office.preview`, `office.diff` | Inspect the proposed result | No |
| `office.verify` | Check the planned result | No |
| `office.commit` | Apply against the source identity and verify replacement | Yes |
| `office.restore` | Optional adapter-backed restore | Yes |

Planned verification does not prove saved output. The commit receipt contains a distinct committed-output verification result.

## Expose JSON tools

Given a session created by your host:

```ts
import {
  AGENT_TOOLS_PROTOCOL,
  AGENT_TOOLS_PROTOCOL_VERSION,
  createAgentToolDispatcher,
} from '@injoffice/agent-tools'

const dispatcher = createAgentToolDispatcher(session, { maxChangeSets: 32 })
const response = await dispatcher.dispatch({
  protocol: AGENT_TOOLS_PROTOCOL,
  protocolVersion: AGENT_TOOLS_PROTOCOL_VERSION,
  requestId: 'capabilities-1',
  method: 'office.capabilities',
  params: {},
})
```

This is an integration fragment: `session` is the session your application creates. An HTTP route, MCP server, or Worker bridge can carry these JSON messages, but the package does not install or authenticate any of those transports for you. Configure bounded retention and call `forget(changeSetId)` when a retained plan is no longer needed.

## Native Office host callbacks

For XLSX, DOCX, and native PPTX, the adapters take host-owned callbacks:

- `snapshot`: identify and project the current authoritative artifact.
- `preview`: apply to an isolated source copy, without mutating stored bytes.
- `apply`: check the expected source identity atomically, apply supported operations, and return the actual replacement.

Use the same native engine behind preview and apply. Never synthesize a successful receipt from the preview model. A persistent host must implement durable compare-and-swap, authorization, and idempotency even if the in-memory session already guards its own lifecycle.

Consult the exact callback types under [agent-office](../reference/generated/packages/agent-office). Formats intentionally have different target and payload schemas.

## Implement another adapter

Implement `AgentArtifactAdapter` to identify, inspect/read, validate, preview, diff, commit, and verify an artifact. Keep artifact and format IDs stable across replacements. Return plain JSON outside the adapter-private artifact value. Advertise only operations you can actually execute and refuse unsupported input before writing.

Abort signals are host-local controls, not wire values. Resource budgets and stable issue codes are part of the integration contract, not optional presentation details.

## Reference contracts

- [Agent-tools session and dispatcher](../reference/generated/packages/agent-tools)
- [Format adapters](../reference/generated/packages/agent-office)
- [Change-set architecture](../reference/generated/contracts/agent-changesets)
- [Optional proposal-host example](../reference/generated/contracts/agent-proposal-host)
