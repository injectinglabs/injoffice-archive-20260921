# Go and the optional HTTP server

Use the Go modules directly when your application owns native file processing. Use `injoffice-server` when an HTTP boundary is useful. Neither is required for the browser-local demo or the documentation site.

## Start the local server

From a source checkout:

```sh
cd go/injoffice-server
go run ./cmd/injoffice-server --addr 127.0.0.1:18765 --artifacts ./artifacts
```

Or run `docker compose up --build` from the repository root. Compose starts the optional server, not the playground or documentation.

## Discover the actual capabilities

```sh
curl -fsS http://127.0.0.1:18765/healthz
curl -fsS http://127.0.0.1:18765/v1/capabilities
```

Liveness is not a full end-to-end document-read/write check. The capability route describes supported routes and mutations, including authentication and persistence boundaries.

## Extract an XLSX file

```sh
curl -sS -D response-headers.txt \
  -H 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' \
  --data-binary @input.xlsx \
  http://127.0.0.1:18765/v1/xlsx/extract
```

This request uploads the file to the server, unlike browser-local WASM. The response includes a native projection and an `X-InjOffice-Artifact-Id` header for the stored artifact. Inspect the HTTP status and error body in your actual client before accepting the result.

`POST /v1/{xlsx,docx,pptx}/extract` supports raw file bodies, multipart uploads, or a stored `artifact_id`. Mutation routes take the documented multipart `original`, `payload`, and `expected_revision` fields, or an artifact ID in place of the original. A JSON schema example is not automatically a complete HTTP request.

Use opaque artifact IDs, never filesystem paths. See the [HTTP server reference](../reference/generated/go/injoffice-server) for the exact route and storage contracts.

## Use Go directly

| Module | Native responsibility |
| --- | --- |
| [xlsxpatch](../reference/generated/go/xlsxpatch) | XLSX extract/apply and bounded chart/pivot/drawing patches |
| [docxpatch](../reference/generated/go/docxpatch) | DOCX extraction and surgical document edits |
| [pptxpatch](../reference/generated/go/pptxpatch) | Native PPTX extraction, mutation, and bounded generation |
| [officecompat](../reference/generated/go/officecompat) | Shared OPC preservation and mutation contracts |
| [slidesqc](../reference/generated/go/slidesqc) | Presentation quality checks |
| [collab](../reference/generated/go/collab) | In-process ordered collaboration hub |

Each module has its own `go.mod`. Run `go test ./...` inside the relevant module rather than assuming there is a root Go module.

## Before exposing a server publicly

The optional server is not a multi-tenant hosted Office service. It does not supply your identity system, SSO, tenant authorization, billing, or production backup policy. Keep its default localhost binding until those boundaries are designed. See the [production checklist](security).
