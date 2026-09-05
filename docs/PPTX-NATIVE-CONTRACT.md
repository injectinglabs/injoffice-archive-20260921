# Native PPTX contract v1

`pptx-native/v1` is InjOffice's stable boundary between authoritative Go OOXML
parsing/patching and pure TypeScript rendering/editing. The normative schema is
[`schemas/pptx-native-v1.schema.json`](../schemas/pptx-native-v1.schema.json). Its
bindings live in [`go/pptxpatch`](../go/pptxpatch) and
[`@injoffice/pptx-native`](../packages/pptx-native).

The authoritative extractor and DOM-free RenderTree now consume this boundary.
Surgical native patching and complete PowerPoint layout remain separate migration
lanes from the architecture ADR.

## Invariants

- `contractVersion` is an exact value, not a best-effort hint. Unknown versions are
  refused before content is interpreted.
- Every document, slide, element, group child, and asset has a durable ID. IDs are
  unique across the document and must not be derived from array position. Slide and
  element arrays retain document order and z-order.
- Geometry uses signed integer EMU for offsets and positive integer EMU for extents.
  Font sizes use hundredths of a point; line widths use EMU; animation distance uses
  integer parts per million. V1 has no floating-point JSON values.
- Every asset, slide, and element declares `provenance`. `parsed` objects require a
  `source` anchor that binds the durable ID to a canonical package part, native
  object ID, and fragment SHA-256. `authored` objects cannot claim source or
  passthrough capabilities. A parsed deck may contain authored objects created by
  a later edit; `sourceRevision` binds all parsed anchors to one immutable revision.
- `passthrough` entries are opaque server-issued tokens. They are not XML, ZIP paths,
  or permission to copy arbitrary bytes. The Go save host resolves a token only
  against the exact source revision, owner part, and fingerprint.
- Parsed source-only assets also require an opaque read capability in their
  `passthrough` array. `asset.source.partName`, digest, byte length, and asset ID are
  integrity/routing metadata, not authority to read or serve a ZIP entry. A gateway
  resolves the capability against the source revision, actual relationship-routed
  media part, full digest, exact byte length, and issuance reason before returning
  bytes; shared media references reuse the same asset and capability.
- Contract validation checks inline asset encoding and declared length. The host that
  accepts or serves asset bytes must additionally verify the declared SHA-256 before
  using them; the synchronous browser validator is intentionally crypto-provider-free.
- Authored decks contain only authored objects. New objects added to a parsed deck
  explicitly use authored provenance and have no anchor until the authoritative
  engine commits and reparses them.
- Unknown JSON properties, unsafe package paths, duplicate IDs, dangling asset or
  diagnostic references, ragged tables, and contradictory compatibility states are
  refused.

## Bounded v1 element model

The element union is intentionally conservative:

- rich-run text boxes and placeholders;
- supported preset shapes with optional rich text, explicit solid/no fill, and
  renderer-neutral outline metadata;
- straight connectors and arrowheads;
- PNG/JPEG-compatible picture assets by stable asset ID;
- strict rectangular plain tables;
- opaque chart frames whose complete relationship closure stays in the Go engine;
- nested groups when a parser can represent the transform and children without
  flattening or loss.

Opaque charts are always `preserveOnly` in v1. The browser may render a referenced
preview asset, but it cannot claim semantic chart editability. Unsupported native
features remain server-side passthrough data and produce compatibility diagnostics;
they are never silently normalized into the closest supported shape.

Chart closures are fail-closed: `opaqueRef.ownerPart` must equal `chartPart`, and the
chart relationship ID must equal the owning element source anchor's relationship ID.
Chart previews must reference image assets. Every parsed slide descendant, including
a group child, must anchor to the exact part named by its owning parsed slide.
Folder conventions are not authority: slide, chart, asset, and passthrough parts are
relationship-resolved and may live outside conventional `ppt/slides`, `ppt/charts`,
or `ppt/media` folders.

Parsed tables use a deliberately narrower exact subset than legacy authored plain
tables. A projected `p:graphicFrame` / `a:tbl` has positive explicit tracks whose
sums equal its frame extents, unmerged rectangular cells, empty table properties,
explicit sRGB/no fill, four explicit no-fill edges, and self-contained native
paragraph/run plus text-body metadata for every cell. `text` remains a validated
newline-joined summary, not a second layout authority. Authoritative and legacy
cells cannot be mixed. Styles, banding, inheritance/theme paint, effects, visible
or omitted borders, merges, rotated/flipped frames, charts, and embedded objects
preserve the exact owning graphic-frame subtree through a capability instead of
producing a partial table.

All package part names are canonical package-relative OPC names with no leading or
empty segment, no `.` or `..` component, and no segment ending in `.`. Uppercase
percent escapes are accepted for legitimate URI bytes, but escapes that decode to
traversal, a trailing dot, `/`, `\\`, or controls are refused before any package
lookup. During extraction, the authoritative Go engine additionally verifies
relationship and content types against the source package; folder names are never a
substitute for those checks.

## Authoritative Go extraction

`pptxpatch.ExtractNativePPTX` is the secure package-bytes entry point for parsed
native v1 decks. It is independent from the older reconstructive `ParsePPTX`
reader. The extractor resolves the root office document and its
presentation/slide/layout/master/theme dependencies through exact OPC
relationships and effective content types; conventional `ppt/` folder names are
never trusted as routing authority. ZIP names are matched by ASCII
case-insensitive, percent-decoded OPC identity while each source anchor and
passthrough request retains the package's actual spelling.

Both Transitional and Strict Office relationship Type namespaces are supported.
Strict changes the Office relationship Type URIs to the purl namespace, but its
`.rels` documents still use the standard OPC package relationships namespace.
Mixing Transitional and Strict presentation, drawing, attribute, or relationship
Type namespaces is refused.

Native extraction is deliberately loss-aware. V1 currently extracts explicit
slide geometry, simple solid backgrounds, and fully self-contained rich text-box
content. Theme latin typefaces (`+mj-lt` / `+mn-lt`, and east-asian/complex-script
tokens that fall back to the same family's latin face) and `schemeClr`
snapshots (`srgbClr` or `sysClr lastClr`, remapped through an exact master
`clrMap`) are materialized into the existing `fontFamily` / color fields.
Documented DrawingML `tint`, `shade`, and `lumMod`/`lumOff` transforms are
applied in document order to that sRGB snapshot. Missing theme slots,
placeholder `phClr`, alpha-only or unmodeled color transforms, and unresolved
tokens remain object-local refusals rather than host-default fonts or
approximated colors. Opaque charts stay preserve-only; a relationship-routed
PNG/JPEG preview is attached as `previewAssetId` and painted at the frame EMU,
while a chart without that exact preview stays a RenderTree refusal rather
than an invented chart renderer. It also extracts exact embedded picture relationships into
source-only assets with full digests and byte lengths. Picture rotation, flips,
crop, effects, links, and other semantics that v1 cannot represent stay
capability-backed `preserveOnly` (or fail closed when relationship integrity is
invalid).

The exact text-body slice materializes DrawingML defaults and direct equivalents
into `textBody`: left/right insets `91440` EMU, top/bottom insets `45720` EMU,
`square` horizontal wrapping, `top` vertical anchoring, fixed `none` autofit, and
fixed `overflow` semantics on both axes. Direct nonnegative signed-32-bit insets,
`wrap="none"`, and `anchor="ctr"`/`"b"` are also exact. Insets must leave a positive
inner rectangle. These fields travel through the canonical schema, Go/TypeScript
bindings, validation, and RenderTree; omitted OOXML defaults and explicit-equivalent
markup produce the same contract value.

Text-body rotation or vertical flow, multiple columns/column spacing, `anchorCtr`,
distributed/justified anchoring, WordArt/warp, font or shape autofit, clip/ellipsis overflow, 3D/effects,
extensions, inheritance/theme dependencies, and unknown body properties are
element-scoped capability-backed refusals. The extractor does not approximate them
with the exact subset. Script, language, and run direction still come from the
injected font-layout host because v1 does not infer them from characters. Literal
tab, CR, or LF characters inside `a:t`, and explicit `a:br` children, require modeled
DrawingML hard-break metrics and therefore produce an exact object-local content
refusal with passthrough for the unchanged shape subtree rather than ordinary glyph
shaping. Other shape-tree, layout, master, theme, relationship, transition,
timing, media, chart, table, and extension content is retained through exact
scoped capabilities and
diagnostics rather than approximated or omitted. If preservation is required and
no trusted token factory is supplied, extraction fails closed.

Nested `p:grpSp` extraction retains source child order and durable object IDs.
`transform` carries DrawingML `off/ext`; the additive `childTransform` carries
`chOff/chExt`. The pair is valid only when it composes to exact integer-PPM scale
and integer-EMU translation, allowing the renderer-neutral affine operation to
scale strokes, shaped text, pictures, and descendants together. Parsed groups do
not use their extent as a clip. Rotated, flipped, fractional, empty, or otherwise
unmodeled or hidden-descendant groups emit no partial native subtree: the exact raw `p:grpSp` remains
bound to slide passthrough and the slide becomes `preserveOnly`. Malformed group
structure fails extraction. Every parsed group requires `childTransform`; authored
v1 groups may omit it and retain the original identity-local group behavior.
If any descendant is itself refused, the extractor discards every staged child,
asset, diagnostic, and capability and preserves the exact owning `p:grpSp` as one
opaque unit; a renderer never receives a partial projection of that group.
Passthrough requests are staged once in bounded extraction-local storage. If any
group is present, the host must supply a `NativePassthroughTransactionalTokenFactory`:
tokens remain unpublished until the complete deck, substituted tokens, and
canonical JSON have validated and one atomic commit succeeds. Refusal, collision,
late validation, or commit failure rolls back every staged capability.

The editable AutoShape checkpoint is narrower than the authored preset enum on
purpose. It accepts `rect`, `ellipse`, `triangle`, and `diamond` only when the
geometry has an explicit empty adjustment list, the transform has no effective
rotation or flip, fill is explicit `noFill` or one untransformed sRGB `solidFill`,
and the outline is explicit single/centered solid `noFill` or sRGB paint. Native
outlines carry cap, join, solid dash, and (for miter joins) the exact miter limit
through the Go, TypeScript, RenderTree, and paint-command boundaries. Legacy v1
strokes that predate those additive fields remain valid; new authoritative
AutoShape extraction always supplies the complete outline metadata. Line and table
border widths are bounded by DrawingML's `20,116,800` EMU line-width ceiling;
miter percentages stay within DrawingML's signed 32-bit positive-percentage domain.

Custom/adjusted geometry, unsupported presets, gradient/pattern/picture/group or
theme fills, non-solid/compound/inherited lines, effects, 3D, and effective
rotation/flips are never flattened to a nearby preset or color. The extractor
emits a source-anchored refused shape with a capability for the exact raw object.
For that refusal-only case `preset` may be absent, and RenderTree emits a visible
placeholder without consulting geometry or paint. A missing preset on editable or
preserve-only shape content is contract-invalid. Element-scoped refusals aggregate
to the slide/deck while remaining element placeholders; an independent slide-level
refusal still blocks the whole slide.

The XML boundary accepts a deliberately canonical v1 declaration only at the
start of a part: XML 1.0, followed optionally by `encoding="UTF-8"` and then
optionally by `standalone="yes"` or `standalone="no"`. Duplicate, unknown,
reordered, later, or non-XML processing instructions are refused. Broader valid
XML quote and encoding-label variants require a later compatibility change with
fixtures; consumers must not silently broaden this parser rule.

`Previous` may reuse validated source-anchor identities across reorder. Duplicate
anchor keys are rejected, and a prior identity namespace is retained only when at
least one current relationship-routed slide has the same exact part plus `sldId`
anchor. An unrelated, otherwise contract-valid prior deck cannot inject its
document ID. Once continuity is established, only exact matched slide anchors and
elements owned by those matched slide parts retain prior IDs. Asset IDs survive
only when an image relationship from a matched slide reaches the same current OPC
part and the prior anchor's digest, byte length, and effective MIME all match the
current bytes. The document ID itself remains durable after any one slide-anchor
match so legitimate slide additions, deletions, edits, and single-slide revisions
do not rename the deck. New document IDs are seeded from presentation content,
and slide/element IDs are scoped by that document ID, avoiding cross-document and
position-derived collisions.

## Resource budgets

Schema-owned limits are generated into both runtimes. V1 refuses inputs beyond 256
MiB of JSON, one million JSON nodes, 64 levels of JSON/group nesting, 4,096 slides
or assets, 10,000 elements per container, or 100,000 elements total. Text is capped
at 1,048,576 UTF-16 code units per field and 16,777,216 total; tables are capped at
one million cells. Inline assets are capped at 512 MiB declared bytes per asset,
about 85 MiB of base64 per field, and 128 MiB of base64 total.
Authoritative extraction separately caps the cumulative bytes inspected across
unique picture parts. That work counter and its digest cache are package-scoped
and monotonic: rolling back an opaque group does not permit the same media bytes
to be hashed repeatedly while refunding the work budget.

Validation stops descending once a structural budget is exhausted, bounds every
array traversal, streams base64 decoding in Go, and caps reported issues. Limits
are refusal boundaries, not recommendations for normal document size.

## Compatibility and refusal

Every deck, slide, and element carries one status:

- `editable`: rendering and supported edits are native and safe; no refusal
  diagnostics are allowed.
- `preserveOnly`: the object survives surgical save but is not semantically editable;
  at least one warning explains why.
- `refused`: the consumer must not claim a safe render/edit/save path; at least one
  refusal diagnostic explains the blocker.

A parent status must be at least as restrictive as its descendants. Diagnostic codes
are stable machine keys; messages are human detail. Optional diagnostic scopes use
durable slide/element IDs or canonical OPC part names. When a scope supplies both
a slide and element ID, the element must belong to that slide.

## Deterministic JSON

`stringifyNativePptx` and `MarshalNativePPTXJSON` implement the same profile:

1. validate before encoding;
2. omit absent optional fields rather than writing `null`;
3. sort assets by durable ID;
4. preserve every order-bearing array;
5. sort object keys by UTF-16 code unit;
6. emit safe integers in base-10 with no negative zero; and
7. emit compact UTF-8 JSON with standard JSON string escaping.

The shared canonical fixture is encoded by both runtimes byte-for-byte. Canonical
JSON is suitable for cache and revision hashes; the original PPTX package remains
the persistence source of truth.

## Migration from `WireDeck`

`packages/slides/src/wire.ts` remains the generated-deck and legacy gateway contract
during migration. It is not `pptx-native/v1` and must not be relabeled or type-cast:

| `WireDeck` | Native v1 |
| --- | --- |
| optional `cx`/`cy` | required `size` in EMU |
| slide/shape array position or optional `key` | required durable IDs |
| overloaded `WireShape.kind` | discriminated text/shape/connector/picture/table/chart/group union |
| `#rrggbb` colors | normalized uppercase `RRGGBB` colors |
| floating-point points/distance | integer OOXML-native units |
| inline picture/chart payload conventions owned by a host | top-level asset manifest and opaque server tokens |
| no source revision or anchors | revision-bound source anchors and passthrough references |
| unsupported content may be absent | explicit capability-backed refusal placeholders |

Migration is staged:

1. `@injoffice/pptx-authored` compiles strict `DeckSpec`/`WireDeck` JSON through
   the DOM-free `@injoffice/slides/authoring` boundary, assigns SHA-256 durable
   IDs, materializes supported Office defaults, and emits validated authored
   native v1 atomically. Unsupported or inherited wire semantics return issues
   and no deck.
2. The Go parser gains source anchors and emits parsed native v1 without removing its
   existing `Deck` API.
3. The pure RenderTree consumes native v1. `DeckView` and `DeckCanvasView` are
   deprecated legacy preview/editor surfaces; production authored-deck rendering
   migrates to the native-v1 RenderTree and they receive no new fidelity authority.
4. Surgical patch APIs accept native operations against `documentId` and
   `sourceRevision`; the existing ParsePPTX → BuildPPTX path remains explicitly
   reconstructive until removed by a later PR.
5. `WireDeck` is deleted only after gateway and website consumers have migrated and
   compatibility fixtures pass.

Index-based fallback IDs are prohibited because they change after reorder and make
collaboration or surgical persistence target the wrong native object.

## Schema and binding drift

The JSON Schema is the canonical source for version, enums, discriminated JSON
shape, and bounds. `npm run generate:pptx-contract` produces the embedded TypeScript
schema/enums, resource limits, and property/required-field manifests plus their Go
equivalents and a schema SHA-256. TypeScript compile-time assertions compare every
interface to that manifest. Go reflection tests compare every ordinary struct and
the complete discriminated-union field set/common required fields. CI runs the
generator in check mode. Both validators consume the same valid/invalid fixture
corpus and must emit the same canonical bytes for the golden fixture.

Adding a field requires a schema change, regenerated bindings, Go/TypeScript binding
updates, and shared fixtures in one PR. A breaking semantic or field change requires
`pptx-native/v2`; consumers must never guess across versions.
