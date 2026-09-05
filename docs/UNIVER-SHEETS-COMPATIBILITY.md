# Univer Sheets compatibility

This matrix freezes the comparison target at Univer `0.25.1`, the version
currently pinned by InjOffice. "Available in the editor" and "round-trips in a
native Office file" are separate claims; a green menu is not persistence proof.

## Public Univer Sheets surface

`@injoffice/univer-sheets/browser` composes every public Sheets preset enabled
by default. `InjOfficeSheetsFeatureComposition` exposes one key vocabulary and
configuration object for these presets and advanced InjOffice providers.

| Capability | Default | Configuration key |
|---|---:|---|
| Core editing, formulas, number formats | on | mandatory |
| Conditional formatting | on | `conditionalFormatting` |
| Data validation | on | `dataValidation` |
| Drawing and image insertion | on | `drawing` |
| Filters | on | `filter` |
| Find and replace | on | `findReplace` |
| Hyperlinks | on | `hyperlinks` |
| Notes | on | `notes` |
| Sorting | on | `sort` |
| Tables | on | `tables` |
| Threaded comments | on | `threadComments` |

Setting a key to `false` omits both the feature model and its UI plugin. The
corresponding ribbon and context-menu items are therefore absent, not merely
disabled. CSS and English locale data are supplied by the composition package.

The focused collaboration demos intentionally register only the core preset.
They are protocol proofs, not the full workbook editor.

The same composition config covers advanced formulas, charts, collaboration,
connectors, edit history, import/export, outlines, pivots, print, range
preprocessing, shapes, and sparklines. They can each be disabled, activated
with UI hidden, or lazy-loaded. Dependencies and required server capabilities
fail before payload activation; failed activations roll back. The provider
bridge supplies a live `menuHidden$` used by the chart, pivot, shape, sparkline,
connector, outline, and print adapters, allowing independent UI visibility while
their commands/models stay active. Providers without live visibility support
refuse that transition instead of changing registry state only.

OSS presets can be omitted independently through the same config, but cannot be
hot-swapped after construction: Univer does not expose a safe public plugin
unregistration API, and each upstream preset couples its model and UI plugins.

## Insert menu

Univer's public drawing, image, hyperlink, and comment Insert commands come
from the presets above. InjOffice's independently implemented chart, pivot
table, and shape actions are added by `registerInjOfficeInsertMenu`. Each is
optional: omit its callback or set its `InjOfficeInsertFeatureConfig` key to
`false` and the command is not registered or shown. `UniverEditor` exposes the
same choice through its `insertFeatures` prop.

## Univer Pro comparison

InjOffice independently covers several areas that Univer categorizes as Pro,
but it does not yet claim behavioral or file-format parity.

| Pro category | InjOffice status | Principal gap |
|---|---|---|
| Collaboration | partial | Client permissions, live-share viewports, a durable outbound reconnect journal, and causal remote/offline undo-redo coordination are available; server authorization/durability, persistent undo history, complete offline conflicts, structure-versus-content-range/object transforms, a concrete Univer command-stack adapter, and complete UI remain |
| Edit history | partial | Lifecycle/model, configurable Univer commands, an accessible timeline, and a fail-closed collaboration restore coordinator are available; automatic Sheets mounting and concrete production storage/authz/atomic reset services remain |
| Import/export | partial | Snapshot/server-unit jobs plus server-authorized coordinators for quiesced-head import and immutable, exact-head export are covered; production transaction/receipt services, the XLSX exporter, arbitrary snapshot serialization, bundled UI, remote acquisition, and unrestricted format coverage remain |
| Printing | partial | An accessible configurable React workflow, optional Univer ribbon command, host-injected preview sessions, bounded native hydration, persisted layout/render commands with atomic snapshot undo, and a permission-aware server-first full-snapshot collaboration protocol are available; automatic portal mounting, a bundled sheet paginator/renderer, durable/offline collaboration service, collaborative undo rebase, repeated-title write-back, complete watermark/custom-paper output, and native conformance remain |
| Charts | partial | All 30 public plot families have renderer-neutral/ECharts mappings, configurable lifecycle/layer commands with snapshot undo, command-routed panel editing, built-in layer controls, host-injected cancellable image export with accessible delivery UI, stable native update/delete for the writable subset, and a bounded server-ordered per-object collaboration protocol; durable/offline transport, collaborative undo rebase, complete styling/effects, renderer conformance, and broader native fidelity remain gaps |
| Pivot tables | partial | Configurable lifecycle/field commands with snapshot undo, command-routed panel editing, ordered stable-ID object collaboration, stable native update/delete, page selections, hidden-member filters, and label sorts are covered; durable/offline transport, collaborative undo rebase, value/label conditions, broader sources, rich layout/formatting, slicers, and timelines remain |
| Sparklines | partial | Native x14 persistence, model, deterministic renderer, configurable Univer lifecycle commands, undo, and ordered server-first object collaboration are implemented; a bundled collaboration service/reconnect queue, collaborative undo rebase, cell renderer, and editing UI remain |
| Outlines/grouping | partial | Row/column model, native XLSX persistence, configurable Univer commands/undo, and an authority-gated server-first collaboration protocol for stable-ID create/update/collapse/remove are available; the outline gutter, hosted durable/offline service, structural-edit coordination, collaborative undo rebase, property-level merge, and native transaction remain |
| Shapes | partial | Stable native update/delete, configurable create/update/remove commands with snapshot undo, and ordered stable-ID collaboration are covered for top-level shapes; durable/offline transport, collaborative undo rebase, grouping, layering, rotation, rich text, gradients, direct-editor transforms, and complete connector authoring remain |
| In-cell graphics | partial | Public drawing preset supplies images; broader Pro graphics are not matched |
| Data connectors | implemented core | Lifecycle includes host authorization, scheduling, cancellation, memory caching, schemas, preprocessing, atomic lifecycle/refresh undo, configurable Univer controls, credential-free round trips through an InjOffice custom OPC part, and server-owned ordered refresh with model/cell conflict checks; production credentials/egress, durable shared caches/logs/leases, Excel-native connection mapping, and complete UI remain host/integration work |
| Range preprocessing | implemented core | Ordered host stages, cancellation, validation, lifecycle errors, connector wiring, SHA-256 collaboration fingerprints, and portable strict manifests are available; direct worksheet/formula-engine adapters, worker/cache policy, and production differential qualification remain |
| Enhanced formula engine | partial | A configurable client facade, 142-function real-engine gate, and ordered authority/revision/spill-safe result sharing are available; Pro performance, complete semantics, authority election, worker behavior, and debugging parity remain |
| Server-side calculation | partial | Deterministic host-injected jobs plus ordered collaborative result distribution and spill refusal exist; no bundled Excel-equivalent engine, managed authority/service, durable result transport, or differential conformance proof |

Parity work must be based on public behavior, public APIs, standards, and
independently authored tests. Proprietary `@univerjs-pro/*` implementation code
is not an InjOffice source dependency.
