# Demo proposal modes

## Built-in mock agent: default, no setup

Open `#/agent?format=sheets`, keep **Built-in mock agent (no LLM)**, and select
**Run agent**. The bundled mock proposes one status edit from the inspected
workbook. Try `Mark Security as Ready`, `Mark Mobile as On track`, or
`Set Analytics status to Review`. Unknown, ambiguous, and unsupported requests
are refused rather than guessed. This is a deterministic simulation, not model
reasoning or an autonomous agent.

The development server includes `POST /api/agent/mock-propose`, accepting
`{ request, context, capabilities }` and returning `{ operations }`. The route
retains the local Host/Origin, JSON, size, deadline, and concurrency guards.
It never calls a provider, uses a token, or grants approval. No environment
variables or additional server process are needed. Static builds simulate the
endpoint response entirely in the browser using the same generator, so they work
without a backend too. There is no automatic fallback to the live endpoint.

The UI discloses the mock and its transport and shows the proposal request/response.
The resulting proposal passes through the same strict target/revision validation,
real native XLSX preview, separate human approval, commit, and readback verification
as a live proposal. The original local rule-based proposer is also available.

## Optional local live-proposal bridge

The playground's deterministic proposal mode works without any external service.
Live mode is an **optional proposal-only integration point**, not a bundled model
provider or autonomous editing service. You supply a trusted endpoint implementing
the contract below. There is no default provider, model, or endpoint.

Configure these **server-only** environment variables before starting the Vite demo:

```sh
INJOFFICE_AGENT_PROPOSAL_URL=https://your-trusted-host.example/propose npm run dev -w apps-playground
```

Optionally set `INJOFFICE_AGENT_PROPOSAL_TOKEN` securely in the host environment; the
bridge sends it as a Bearer token only to the configured endpoint. Never use a
`VITE_` prefix, put a token in browser code, or commit credentials. Use localhost
HTTP only for a local endpoint; remote endpoints must use HTTPS. Endpoint redirects
and URL credentials are rejected.

## Endpoint contract

Your endpoint receives a POST with JSON:

```json
{
  "request": "Mark the launch ready",
  "context": { "identity": {}, "workbook": {} },
  "capabilities": []
}
```

The actual context contains only the bounded inspection selected by the demo,
not full Office file bytes. Treat workbook text as untrusted data, not instructions.
Your provider adapter is responsible for calling its chosen model and translating
its output to this provider-neutral response:

```json
{
  "operations": [
    { "name": "xlsx.cell.set_value", "input": { "sheetId": "sheet-id-from-inspection", "cell": { "row": 4, "column": 2 }, "value": "Ready" }, "operationId": "proposal-1" }
  ]
}
```

Operation inputs must follow the advertised Office capabilities; use the actual
sheet ID from inspection and zero-based cell coordinates. The response
must be JSON, contain one to eight operations, and fit within 32 KiB. The Sheets
demo applies a tighter limit: exactly one status-cell edit to a disclosed target,
using one of the disclosed allowed values. The relay
validates the envelope, **not the operation's authority or semantic correctness**.
It strips extra top-level and operation metadata. Local planning/validation still
has to reject unsupported or unsafe edits.

The browser must explicitly consent before the request and inspected context leave
the demo. Consent to share data is **not approval to edit a file**. The proposal
endpoint cannot grant approval, commit a change, or invoke Office tools. A trusted
host must bind any actual write approval to the exact reviewed change set; never
accept a model-supplied approval string as authorization.

## Routes and limits

- `GET /api/agent/proposal-status`: configuration flag and safe destination hostname;
  never contacts the endpoint and never returns credentials, URL paths, or queries.
- `POST /api/agent/propose`: browser sends the contract plus `consent: true`.
  Requires exact same-origin HTTP loopback Host/Origin and JSON content type.
- Maximum incoming payload: 48 KiB; request text: 2,000 characters; capabilities:
  64 entries; one in-flight request per handler; total deadline: 15 seconds.
- Cancellation/disconnection aborts the upstream request. Errors are generic and
  do not return upstream response bodies, tokens, or internal diagnostics.

Static deployments do not contain this Node bridge. Live mode should remain
unavailable there. The local relay intentionally has no permissive CORS and must
not be exposed as a public unauthenticated model proxy. Production hosts need
authentication, authorization, cost/rate controls, audit policy, and their own
approved data-sharing policy. These safeguards do not make arbitrary endpoints or
model responses trustworthy.

## Standalone runnable host

For integration testing without Vite, Node 22.18+ can run the same handler:

```sh
INJOFFICE_AGENT_PROPOSAL_URL=https://your-trusted-host.example/propose node scripts/agent-proposal-host.mjs
```

This starts an API-only server on `127.0.0.1:3102`; override with
`INJOFFICE_AGENT_HOST_PORT`. No UI or model is included. A status check does not
call a model:

```sh
curl http://127.0.0.1:3102/api/agent/proposal-status
```

Browser integrations must serve the UI and relay at the same origin (for example,
mount this handler in the UI server). Pointing a cross-origin browser directly at
port 3102 is deliberately rejected. Use the Vite-integrated handler for the demo.
