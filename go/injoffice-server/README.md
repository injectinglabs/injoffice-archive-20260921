# injoffice-server

Optional local HTTP server for native XLSX, DOCX, and PPTX extract and
mutations plus the collaboration protocol in `docs/COLLABORATION-PROTOCOL.md`.
XLSX reuses the same handlers as `go/xlsxpatch/cmd/xlsxnative`
(`ExtractNativeWorkbookV2` and `ApplyNativeWorkbookMutationPayloadV1`). DOCX
and PPTX call `ExtractNativeDocumentV1` / `ApplyNativeTextMutationPayloadV1`
and `ExtractNativePPTX` / `ApplyNativePPTXMutationPayload`. Packages are stored
under opaque artifact IDs. Join/presence/ops fan out through `go/collab` over
HTTP + SSE. There is no Injecting auth, agents, SSO, billing, or tenant RBAC.

```bash
cd go/injoffice-server
go run ./cmd/injoffice-server --addr 127.0.0.1:18765 --artifacts ./artifacts
```

Bind address is configurable (`--addr` / `INJOFFICE_ADDR`). The default is
localhost. Docker Compose publishes `127.0.0.1:18765` on the host; the
container itself listens on `0.0.0.0:18765` so the published port works.

From the repository root, Compose publishes localhost:18765 and keeps artifacts on a named volume:

```bash
docker compose up --build
```

`GET /healthz` returns a lightweight process-liveness response. `GET
/v1/capabilities` returns the versioned route and mutation-operation inventory,
including the server's authentication and collaboration-persistence limits.

```bash
curl -sS -H 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' \
  --data-binary @input.xlsx http://127.0.0.1:18765/v1/xlsx/extract
```

`POST /v1/{xlsx,docx,pptx}/extract` accepts a raw OOXML body, a multipart file,
or JSON `{"artifact_id":"..."}`. Uploading bytes mints an opaque id returned as
`X-InjOffice-Artifact-Id`. `POST /v1/{xlsx,docx,pptx}/mutations` accepts the
helper multipart (`original`, `payload`, `expected_revision`) or `artifact_id`
in place of `original`; a stored mutation is written back under the same id.

Artifact IDs are minted tokens such as `art_<32 hex>`. They are not workspace
paths and are not used as filesystem path components.

Collaboration RPCs take that same opaque id as `path`. Mint a live connection
with `POST /v1/collab/session`, stream `collab.*` frames from
`GET /v1/collab/events?session_id=...` (SSE), then:

| RPC | HTTP |
|---|---|
| `collab.join` | `POST /v1/collab/join` |
| `collab.leave` | `POST /v1/collab/leave` |
| `collab.presence` | `POST /v1/collab/presence` |
| `collab.op.submit` | `POST /v1/collab/op/submit` — `409` + `{"error":"STALE_BASE"}` when `base_seq` is not head |
| `collab.op.since` | `POST /v1/collab/op/since` |

Two browsers can share one artifact from the playground: start this server,
then `npm run dev` from the repository root and open `#/collab` twice.
