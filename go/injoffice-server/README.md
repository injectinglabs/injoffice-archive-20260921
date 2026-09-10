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

## Opt-in native DOCX pages

Native DOCX page preview is disabled by default. To enable it locally, build
the workspace packages and `@injoffice/docx-page-paint-worker` using Node 22,
then pass the absolute compiled worker path:

```sh
go run ./cmd/injoffice-server -docx-preview-worker /absolute/injoffice/apps/docx-page-paint-worker/dist/worker.js
```

Set the playground's existing `VITE_INJOFFICE_API_BASE` to this helper. Opening
a DOCX still does not upload it: the separate **Upload to helper and render
native pages** button names the destination and explicitly submits the current
bytes to `POST /v1/docx/page-preview`. The endpoint does not persist the upload.
The SVG viewer validates the returned page-paint schema and source package
digest, mounts one page at a time, and clears stale pages after edits/reopen.
It is a read-only glyph preview, not a selectable Word editor.

The server reconstructs layout, settings, embedded font assets and referenced
PNG bytes from the submitted archive. The pinned HarfBuzz worker shapes,
paginates, outlines and validates native page paint. Missing fonts, unsupported
layout and unavailable providers refuse rendering; the separately labeled
approximate content preview remains available without weakening edit safety.

This route limits packages to 8 MiB, permits one compilation at a time, checks
cancellation between extraction passes, and limits the subprocess to 30 seconds
(within a 45-second request context), 64 MiB framed output and a 512 MiB V8 heap.
These are not OS-level total-memory isolation: WebAssembly and Go allocations
are separate. The browser streams at most 16 MiB and bounds displayed geometry.
Keep the default loopback listener. This helper has no authentication and is
not a public multi-tenant rendering service.
