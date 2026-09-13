# @injoffice/pptx-native

`NativePptxTableInspection` is a separate read-only source inspection contract.
`decodeNativePptxTableInspection(value, strictDeck, packageSHA256)` validates it
against the caller's strict extraction and independently computed package hash.
Its plain paragraphs and stored rectangles carry no native-deck capabilities
and omit authored table borders, fills and source font styling. Use the browser
`@injoffice/pptx-wasm` client's `inspectTables` method to own the byte snapshot,
hashing and strict extraction automatically. The decoder does not parse ZIP/XML;
opaque object IDs and whole-part hashes remain evidence from the native producer.

Text/shape transforms may include `quarterTurns: 1 | 2 | 3`, clockwise around
their frame center. Omission means zero rotation. The renderer composes exact
integer-affine quarter turns; 90/270-degree frames require matching width/height
parity to avoid fractional EMU centers. Group, picture, table, and connector
rotation is not introduced by this field. Parsed rotated shapes remain
preview-only, and mutation requests cannot supply quarter-turn transforms.

Text extraction resolves a bounded local DrawingML style cascade: list default,
matching list level, paragraph properties, then explicit run properties. Only
modeled typeface/size/bold/italic/solid color and paragraph alignment/list metadata
participate. Theme tokens still resolve through the relationship-bound theme.
Source XML bytes are never rewritten to materialize these rendering properties.

For non-placeholder AutoShapes, an authored `fontRef` can supply a missing Latin
font and/or solid color after that local cascade. The bounded projection requires
graphic ASCII text and an exact relationship-bound theme. Applicable presentation
and master text-style layers must be absent or empty; competing defaults are not
guessed. Explicit local/list font and color retain precedence. No missing size,
bold, italic or paragraph values are invented. These shapes remain read-only,
with `pptx.shape-font-reference-preview` identifying the source projection. This
is a prerequisite for broader text inheritance, not full slide fidelity.

Paragraphs may retain an authored `bulletCharacter` (one Unicode scalar, with
`bullet: true`) and bounded `marginLeftEmu` / `indentEmu`. The approximate file
preview uses those values instead of inventing a generic bullet. Exact native
RenderTree still refuses marker/indent geometry until qualified. Auto-numbering,
bullet font overrides and unmodeled spacing are not silently flattened.

Pictures may carry optional `crop` with all four source-edge insets (`left`,
`top`, `right`, `bottom`) in DrawingML 1/1000-percent units. Each is an integer
from 0 through 99999 and opposing sums must be below 100000. Omitted crop means
the full source image. The native extractor now qualifies positive source crops;
negative/outset and empty rectangles remain preserve-only. Cropping does not
rewrite image bytes or grant any new image mutation capability.

The dependency-free, versioned native PPTX JSON contract shared by InjOffice's Go
parser/patcher and future pure TypeScript RenderTree. It contains no renderer,
React, Konva, DOM, ZIP, storage, or IPC code.

```ts
import { assertNativePptx, stringifyNativePptx } from '@injoffice/pptx-native'

assertNativePptx(value)
const canonical = stringifyNativePptx(value)
```

The normative schema is `schemas/pptx-native-v1.schema.json` in the repository.
Generated schema constants, enums, resource limits, and object-shape manifests carry
its SHA-256. Compile/reflection tests fail when TypeScript or Go fields and
requiredness drift from the schema.

Native shape outlines can carry additive cap/join/dash/miter metadata end to end.
Authoritative AutoShape extraction supplies the complete set. A refused shape may
omit its preset so renderers can show a durable-ID placeholder without claiming a
fabricated geometry; non-refused shapes still require a supported preset.
Generated resource constants also pin DrawingML's line-width and percentage bounds,
so Go and TypeScript cannot accept different hostile numeric ranges.

Text boxes and text-bearing shapes can carry the exact v1 `textBody` slice: bounded
integer EMU insets, square/no horizontal wrap, top/center/bottom vertical anchor,
and fixed no-autofit/overflow semantics. The authoritative extractor materializes
OOXML defaults into explicit values. Theme latin tokens and scheme
colors, including documented tint/shade/lumMod transforms, are materialized into `fontFamily` and run/shape color fields when the
relationship-routed theme supplies exact snapshots; missing font bytes, empty
slots, and alpha-only or unmodeled color transforms remain refusals. Unsupported body layout is not
normalized to this subset; it is an element-scoped capability-backed refusal.

See [`docs/PPTX-NATIVE-CONTRACT.md`](../../docs/PPTX-NATIVE-CONTRACT.md) for
identity, passthrough, compatibility, canonical JSON, and WireDeck migration rules.

### Source-positioned table inspection

`createNativePptxTableGeometryPreview(inspection, 'host-sans-12pt-clipped-v1')`
creates a read-only arrangement plan from the immutable result of
`decodeNativePptxTableInspection` (including `inspectTables` in the browser WASM
client). Cloned or deserialized inputs must be decoded against their source deck
and package hash again. The planner does not accept arbitrary geometry as source
authority and exposes no mutation or passthrough tokens.

Coordinates use 96 CSS pixels per inch. Each slide's extent is the union of its
inspected table frames, with its original EMU origin retained; this is not a full
slide rectangle. Positions, overlaps and cell geometry are retained. A whole
arrangement exceeding 16384 CSS pixels in either dimension is omitted with a
reason, without fitting or relocating individual tables. Inspection omissions
remain present.

The named host display policy is generic sans-serif 12 pt (16 CSS pixels), 20 px
line height, 2 px inset, top-left wrapping and clipping within each stored cell.
Hosts must disclose that fonts, styles, borders, fills, alignment and spacing are
not authored Office paint. Keep the full reading-order text available because
clipping can hide text. The playground exposes this through a separate explicit
button after browser-local table inspection, without changing native rendering.

An additional explicit `{policy: 'source-no-style-solid-border-v1'}` argument
requests qualified source paint in the arrangement plan. Eligibility requires
one unmerged cell, the actual embedded style's whole-table no-fill/no-border
structure, no conditional style regions, and four equal explicit no-fill or
solid centered single-line borders. Solid lines must have flat caps, round joins,
no arrows, and widths at most 10 pt. The style identifier alone grants nothing.
Theme colors use the source master mapping; unsupported slide/layout overrides
prevent paint qualification. Style, theme, master and layout part hashes remain
in the evidence. The planner's solid border is one centered rectangle with round
joins; this bounded replay policy is not a PowerPoint corner-fidelity claim.

The playground's separate source-paint button applies eligible borders and removes
inspection guides for eligible no-border tables. Other tables retain guides and
visible paint omissions. Transparent fills expose the host canvas. Text still uses
the original approximate host font/layout policy, and the strict native renderer,
source bytes and mutation permissions remain unchanged.


Preset table borders require the additional `source-no-style-preset-border-v1`
paint option. The existing `source-no-style-solid-border-v1` option continues to
omit dashed borders. Both retain the explicit `host-sans-12pt-clipped-v1` text
policy and the same one-cell, structural no-style, no-fill, equal-edge source
qualification. Only explicit flat caps, centered single strokes, round joins,
no arrowheads and widths up to 10 pt qualify; custom dashes and unequal edge
patterns remain omitted. Native table rendering is unchanged.

The preset option accepts the ten non-solid values defined by ECMA-376 Part 1
(2016), §20.1.10.49, and converts their alternating painted/gap runs into CSS
lengths using the source line width. See the [primary ECMA-376 standard download](https://ecma-international.org/publications-and-standards/standards/ecma-376/).
Its returned `dashArray` is immutable. The demo replays a closed SVG rectangle
clockwise from the upper-left corner with zero dash phase. These repeat lengths
come from the standard; PowerPoint's per-edge phase and corner placement have
not been independently qualified. This is an explicit preview policy and does
not establish Office raster parity or restore authored typography.
