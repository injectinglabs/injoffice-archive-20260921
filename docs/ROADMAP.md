# InjOffice roadmap

This roadmap describes the public libraries in this repository. It does not track any particular product deployment, authentication system, or hosted service.

Compatibility claims are scoped to documented and tested public behavior;
partial coverage must never be described as complete.

## Available foundations

### Spreadsheets

- Native file API: `xlsxpatch.ExtractNativeWorkbookV1` / `ExtractNativeWorkbookV2` → `@injoffice/sheets` JSON → TypeScript sheet paint (preview) → mutation JSON → Go `Apply*`. Univer floating UI is an optional editor shell, not file authority.
- ChartSpec with pure data extraction, ECharts option generation, Univer floating UI, configurable lifecycle and per-sheet layering commands, snapshot undo/redo, host-injected cancellable image export, server-ordered per-object collaboration with object/layer revision guards and authority hooks, XLSX reading/writing, stable native identity, surgical update/delete, and grid anchors.
- PivotSpec with deterministic aggregation, page/member filtering, label
  sorting, editing UI, configurable Univer lifecycle/field commands, snapshot
  undo/redo, ordered server-first object collaboration, native XLSX writing,
  stable native identity, surgical update/delete, native pivot/cache hydration,
  and fail-closed conversion of unsupported definitions.
- ShapeSpec with a broad preset catalogue, SVG geometry, Univer floating UI,
  configurable lifecycle commands, snapshot undo/redo, DrawingML
  reading/writing, and ordered server-first object collaboration for the
  supported lifecycle.
- Validated row/column outline groups with nested-range rules, collapse state,
  structural transforms, fail-closed native XLSX metadata persistence, and
  optional configurable Univer command/undo integration; plus authority-gated,
  server-first collaboration with stable IDs, SHA-256 snapshot preconditions,
  room-bound resync, and ordered create/update/collapse/remove operations.
- Renderer-neutral line, column, and win/loss sparkline models with validation,
  deterministic geometry, grouping, host-owned JSON snapshots, live handles,
  and optional configurable Univer command/undo integration.
- Data connectors with host-injected fetching and authorization, bound-range
  refresh, interval/on-open scheduling, cancellation, memory TTL caching,
  positional schemas, atomic lifecycle/refresh cell snapshots, optional
  configurable Univer commands/UI, and a configurable preprocessing pipeline
  with deterministic peer contracts; credential-free definitions also
  round-trip through an explicitly versioned custom OPC part with fail-closed
  graph validation and byte-preservation of unrelated package content, while
  server-owned ordered collaborative refresh uses model and cell SHA-256
  preconditions.
- A host-injected durable edit-history lifecycle with author attribution,
  validated filtering and pagination, isolated historical previews,
  restore-as-new-version lineage, retention metadata, serialized workspace
  workflows, configurable public Univer commands, an optional accessible React
  timeline, a fail-closed collaboration reset/rejoin coordinator, and grid/text
  diffs.
- A curated 142-function formula compatibility audit executed through a
  host-facing Univer worksheet/formula facade, with deterministic custom
  function registration/unregistration, transactional rollback, and
  capability/config gating; plus a server-friendly host-injected calculation
  job protocol with deterministic request fingerprints, typed scalar/error/spill
  results, cancellation, timeouts, lifecycle events, engine negotiation, and
  collaboration-safe revision/fingerprint guards; plus authority-epoch result
  distribution with ordered apply and explicit spill-conflict refusal.
- Print areas, page setup, margins, and header/footer token helpers.
- A renderer-neutral print facade with persisted, validated layout/render
  configuration commands, host-atomic saves, snapshot undo/redo, a validated
  and cancellable host-injected page-preview lifecycle, screenshot/clipboard
  adapters, fail-closed native conversion and hydration, and an accessible,
  responsive, host-rendered React print workspace with configurable controls
  and an optional Univer Start-ribbon entry; plus permission-aware,
  server-first ordered collaboration for the complete persisted configuration,
  with SHA-256 conflict guards, explicit resynchronization, and collaborative
  command/undo routing.
- Host-neutral XLSX snapshot and server-unit import/export jobs with monotonic
  progress, cancellation, structured errors, safe load/replace and save-as
  boundaries, storage/transport injection, and adapters for the existing native
  WASM extractor and optional server artifact store; plus a server-authorized
  collaborative import coordinator with quiesced-head replacement checks,
  atomic commit receipts, pre-commit resume/cleanup, and explicit post-commit
  activation recovery; plus a collaborative export coordinator that authorizes
  and quiesces an exact fully saved room head, exports only an immutable captured
  server unit, binds the output revision/length/SHA-256 to that head, and blocks
  uncertain outcomes until request-ID resolution.
- One deterministic configuration/snapshot vocabulary for the complete public
  Univer Sheets preset set and every advanced InjOffice Sheets domain, with a
  lazy disposable provider bridge, dependency and server-capability checks,
  rollback, host-owned payload splitting, and observable live menu visibility
  in the chart, pivot, shape, sparkline, connector, outline, and print adapters.
  Upstream OSS presets remain construction-time choices because Univer exposes
  no safe public plugin-unregistration API.

The xlsxpatch Go module applies explicit OPC part changes against original workbook bytes and verifies every untouched part. Its preservation inventory additionally guards rebuilt replacements by fingerprinting charts, pivots, drawings—including unknown objects—media, controls, embeddings, relationships, and workbook/worksheet references.

### Documents

- Native file API: `docxpatch.ExtractNativeDocumentV1` → `@injoffice/docs` JSON → TypeScript page paint (preview) → mutation JSON → Go `Apply*`.
- Block and line-aware pagination.
- Header/footer token expansion and comment-thread helpers.
- Surgical DOCX paragraph edits with refusal for unsafe non-text paragraphs.
- DOCX images, charts, notes, themes, and bounded text-watermark support.

### Presentations

- Native file API: `pptxpatch.ExtractNativePPTX` → `@injoffice/pptx-native` JSON → `@injoffice/pptx-render` paint (preview) → mutation JSON → Go `Apply*`. `@injoffice/pptx-authored` compiles `DeckSpec`/`WireDeck` into that native contract.
- Plain-JSON DeckSpec, editing helpers, themes, diagrams, transitions, and deterministic quality checks. React DOM and canvas views in `@injoffice/slides` are legacy preview surfaces, not Office rendering authority.
- PPTX reading/writing for the documented bounded model, including text, shapes, connectors, tables, pictures, charts, transitions, and selected animations.
- Equivalent TypeScript and Go quality-check implementations.

### Collaboration

- Presence models for sheets, documents, and decks.
- A host-supplied transport boundary.
- Linear operation-log synchronization, stale-base catch-up, cell reconciliation, and structural row/column transforms.
- Revisioned, deny-by-default client capability projections with host mutation
  enforcement hooks and atomic role updates.
- Presenter/follower live-share sessions with validated worksheet viewports,
  ordered-event guards, permission revocation, resync, and disposal lifecycle.
- A versioned host-storage outbound journal with stable idempotency keys,
  ordered acknowledgement/replay, resync/rebase boundaries, injected
  exponential retry scheduling, cancellation, and corruption refusal.
- A bounded collaborative undo/journal coordinator that durably stages
  inverse/redo batches until atomic host execution succeeds and distinguishes
  causally earlier remote writes from later ownership changes.
- A documented server protocol with no dependency on a particular backend.

### PDF

See [ROADMAP-PDF.md](ROADMAP-PDF.md).

## Current compatibility limits

These are API limits, not hidden implementation promises:

- Native pivot hydration covers direct worksheet-range sources and the aggregation vocabulary supported by PivotSpec. Page fields, named/external sources, slicers, timelines, and unsupported aggregations are reported and refused by fail-closed conversion.
- Native pivot update/delete resolves the full table/cache relationship graph, retains table-part identity on update, preserves shared caches, and refuses stale, malformed, ambiguous, or unsupported graphs.
- All 30 chart families in the public Univer Charts guide have renderer-neutral
  ECharts mappings. Native XLSX reconstruction remains limited to the explicit
  writer subset and does not preserve every custom style, effect, extension,
  or unsupported OOXML plot type. Exact archive preservation remains safe when
  the part is untouched.
- PPTX ParsePPTX followed by BuildPPTX preserves only the documented model. It is not a general unknown-part round trip.
- Collaboration transforms row/column insertion, removal, and block moves. A host-neutral two-phase undo/redo planner rebases cell/range inverses and selections, neutralizes remotely overwritten cells (including content-only move/reorder mutations), rejects stale plans, and fails closed on ambiguous structure or unsupported objects. Its durable-journal coordinator stages inverse/redo operations until atomic host execution commits and causally protects still-outbound local ownership. Structure-versus-content-range transforms, persistent undo history, and a concrete Univer command-stack adapter remain incomplete.
- Collaboration permission and live-share models are client-side projections.
  Production identity verification, server authorization, ordered event
  assignment/replay, durability, and user-facing permission/presentation UI
  remain host responsibilities.
- The outbound journal and undo coordinator recover pending client submissions,
  but not the in-memory undo stack itself. Their injected storage and transport
  do not establish server durability, authentication,
  encryption, quota policy, or complete offline conflict semantics. Hosts must
  keep client IDs stable and deduplicate `(room, clientId, idempotencyKey)`.
- Spreadsheet connectors deliberately do not implement credential storage,
  unrestricted URL fetching, or authorization. Undoable command snapshots
  reject URL userinfo, queries, and fragments; hosts must supply a stable,
  credential-free gateway URL and keep secrets in their injected fetcher. The
  custom connector OPC extension preserves that same bounded definition subset;
  it is not an Excel Power Query/external-connection mapping and has no
  production Office interoperability certification. Collaborative mode
  delegates authenticated fetching, ownership leases, durable ordering, and
  atomic workbook commits to the host server.
- Spreadsheet outline header controls are not yet implemented. Native XLSX
  persistence covers nested/disjoint groups, collapsed state, summary direction,
  and exact preservation of unrelated package parts, with the format's seven
  persisted nesting levels enforced at the file boundary. Outline collaboration
  does not bundle authenticated durable transport, offline replay, presence,
  structural-edit coordination, property-level merging, rebased undo, or an
  atomic native-file transaction.
- Sparkline models have fail-closed x14 XLSX persistence, optional Univer
  lifecycle commands, and a server-first ordered object-operation protocol
  with fingerprint conflict checks. There is no bundled collaboration server,
  reconnect journal, collaborative undo rebase, automatic cell renderer, or
  editing panel.
- A bundled concrete print-preview renderer, automatic Sheets portal mounting,
  repeated-title native write-back, custom paper dimensions, complete watermark
  rendering/persistence, full header/footer syntax, and complete native
  print-setting conformance remain incomplete. Persisted print configuration
  has a bounded online collaboration protocol, but production authentication,
  durable sequence storage, offline replay, property-level merge, collaborative
  undo rebase, and atomic native-file transactions remain host work.
- Edit history has configurable Univer workflow commands and a host-mountable
  timeline/preview/restore UI, but still needs automatic Sheets sidebar
  integration, a concrete production storage service, permission policy, and
  a concrete atomic collaboration-server reset/rejoin implementation.
- XLSX exchange delegates arbitrary snapshot/server-unit export to a host codec;
  the in-repo native writer still accepts only explicit supported mutation
  batches. Remote URL acquisition, a bundled open/save UI, and all-feature
  repeated round-trip certification remain incomplete. Collaborative import and
  immutable-head export are host/server integration contracts: this repository
  does not provide their production authentication/authorization, durable
  transaction and receipt stores, operation-log reset/quiescence service,
  immutable capture lifecycle, reconnect service, destination persistence, or
  production XLSX exporter.
- Server calculation is an engine-neutral orchestration contract, not an Excel-
  compatible evaluator or hosted service. A production engine/transport,
  authority election, dependency graph, durable result transport/job store,
  volatile-function policy, performance qualification, and differential
  conformance corpus remain.
- The client formula bridge still delegates dependency evaluation, arrays and
  spills, volatile semantics, worker execution, performance, and diagnostics to
  pinned Univer OSS. Shared derived results have revision/authority/spill guards,
  but the 142-function gate remains curated coverage rather than complete Excel
  or Univer Pro semantic/performance parity.

## Near-term work

1. Expand pivot hydration beyond the bounded page-selection, hidden-member-filter, and label-sort subset to value/label conditions, richer layouts, slicers, and timelines while retaining explicit representability results.
2. Expand the bounded stable-identity native lifecycle shared by charts, shapes, and pivots into the general workbook mutation transaction.
3. Add a cell-diff save protocol that can apply editor changes directly to original XLSX bytes, reducing reliance on guarded rebuilt workbooks.
4. Preserve additional chart styling and unsupported DrawingML extensions without reconstruction.
5. Complete structure-versus-content-range collaboration transforms and bridge the public collaborative undo planner into Univer command-stack transactions.
   Chart-object collaboration is currently a bounded online protocol; durable
   transport, offline replay, property-level merge, and transformed local undo
   remain host/application work.
6. Add interoperable fixture corpora generated by Excel, LibreOffice, Google export, and independent OOXML libraries.
7. Stabilize package APIs toward 1.0 with migration notes and compatibility guarantees.
8. Close the explicitly tracked editor-compatibility gaps; do not claim
   blanket parity until durable UI, model, file, collaboration, and server
   evidence supports it.
9. ~~Add fail-closed native XLSX sparkline extraction, surgical write-back, and optional editor commands~~ (implemented with x14 group diagnostics, unrelated-extension preservation, configurable public Univer commands, and undo snapshots); cell rendering and editing UI remain.

## Host-owned work

The following belongs to an embedding product, not this repository:

- Product-specific gateways, dashboards, authentication, tenanting, credentials, and rollout state.
- Production artifact authorization, durable storage, and retention policy beyond the local OSS `injoffice-server`.
- Connector credentials and network egress policy.
- Agent tools, prompts, delivery workflows, and product feature flags.
- Authenticated multi-user acceptance tests for a specific application.

The optional in-repo `injoffice-server` is implemented as a local development sidecar with generic filesystem artifact storage; the WASM path remains experimental and off by default. Both must remain optional and must not import consumer-specific backend, authorization, tenant, or credential code. Hosts still own production authorization, durable storage, tenant isolation, retention, and collaboration-server conformance. Native paint remains preview mode; Univer remains an optional editor shell, not the file authority.
