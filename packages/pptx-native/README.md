# @injoffice/pptx-native

Text extraction resolves a bounded local DrawingML style cascade: list default,
matching list level, paragraph properties, then explicit run properties. Only
modeled typeface/size/bold/italic/solid color and paragraph alignment/list metadata
participate. Theme tokens still resolve through the relationship-bound theme.
Source XML bytes are never rewritten to materialize these rendering properties.

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
