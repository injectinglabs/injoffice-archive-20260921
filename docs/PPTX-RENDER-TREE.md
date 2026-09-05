# Native PPTX RenderTree v2

`@injoffice/pptx-render` is the one native slide rendering core for InjOffice.
It consumes the validated `pptx-native/v1` boundary emitted by the authoritative
Go package engine and returns an immutable, bounded `pptx-render-tree/v2` value.
This is a preview boundary, not a new persistence model: OOXML relationships,
opaque parts, and surgical saves remain authoritative in Go.

## Rendering contract

The tree keeps slide geometry in integer EMU from compilation through paint
recording. It preserves durable source element IDs, array z-order, element/group
transforms and clips, slide background, supported preset paths, exact native
solid fill and outline metadata (width, cap, join, dash, and miter limit), straight
connectors and named arrow flags (including `w`/`sz` sm/med/lg mapped to
integer EMU from stroke width), picture asset references, conservative tables,
paragraph/run structure, and opaque chart preview images. No result depends on a
viewport, display scale, or device pixel ratio.

For parsed native groups, the authoritative extractor retains `off/ext` as the
group transform and `chOff/chExt` as its child coordinate transform. RenderTree
accepts that mapping only when it produces exact integer-PPM scale and integer-EMU
translation, then emits the affine operation so strokes, shaped text, pictures,
and nested descendants inherit the same scale. A DrawingML group extent defines a
coordinate mapping, not a clip. Rotation, flips, fractional mappings, and
unmodeled group markup remain opaque instead of being approximated by a browser
layout or canvas measurement API. Parsed groups without `childTransform` are
invalid. Legacy authored v1 groups that omit `childTransform` retain their original
identity-local translation and clip. Compilation also composes every nested affine
with exact bounded rational `BigInt` arithmetic. Sub-EMU world positions remain
exactly represented by the emitted local integer-PPM transforms; only genuinely
out-of-budget cumulative world bounds are refused before producing paintable
output.

All native compatibility states remain visible:

- `refused` becomes a placeholder and refusal diagnostic;
- `preserveOnly` remains labeled by a stable render diagnostic;
- a chart without a preview becomes a missing-preview placeholder;
- a font-resolution or shaping refusal atomically refuses the text body and
  paints one text placeholder without partial sibling runs; and
- an asset stays an ID/digest reference whose `resolutionSource` tells the paint
  adapter to resolve it from the validated source deck or a host store. RenderTree
  deliberately never embeds or claims to supply the bytes. The adapter must verify
  resolved bytes against the reference `sha256` and `byteLength` before decode or
  paint.

An element-scoped refusal does not make sibling content disappear. Even though
compatibility aggregation makes the owning slide `refused`, the compiler walks
the elements when every refusal diagnostic is element-scoped and emits a stable
placeholder carrying that element's durable ID. A refusal whose scope is the
slide itself still becomes one whole-slide placeholder. Refused AutoShapes may
omit `preset`; the compiler checks compatibility before geometry and never invents
a fallback rectangle.

The recording surface is the test oracle and renderer-neutral integration point.
The Canvas2D command adapter is deliberately generic: the website supplies its
own context and implements command execution. The package never creates a
drawing element or imports UI/runtime frameworks.

## Native text boundary

The compiler calls the `@injoffice/font-metrics/layout` manifest, resolver, load,
and shaper contracts for every run. It uses only returned integer glyph advances
and metrics. Width estimation from text length and host measurement APIs are
prohibited.

Provider results are runtime trust boundaries, not trusted TypeScript values. Each
resolver result, font resource, refusal, face, decision, attempted-face list,
metric set, glyph, cluster, and shaped segment is exact-key validated and rebuilt
field-by-field under resource budgets. Unknown fields, cycles, malformed statuses,
or inconsistent identities fail closed before tree construction.
The compiler snapshots the validated deck plus manifest and provider identity/callables before invocation,
requires every resolved face and explicit content digest to match that manifest, hashes loaded bytes, derives
the only accepted scaled metrics from the loaded design metrics, and caches one
isolated resource per exact face under provider-call and cumulative-byte budgets.

Native PPTX v1 carries an exact bounded text-body slice: integer EMU insets,
`square`/`none` horizontal wrapping, `top`/`center`/`bottom` vertical-anchor
metadata, fixed `none` autofit, and fixed `overflow` semantics on both axes. The compiler
projects the inner body rectangle and wraps only at complete safe UTF-16 clusters
returned by the injected shaper and classified as an allowed boundary from their
actual text by the shared `@injoffice/font-metrics/layout` conservative Office
classifier. NBSP/glue and word joiner prohibit a break; ZWSP and hyphen/slash add
explicit opportunities; Han/Hangul plus common East-Asian opening/closing punctuation follows bounded
line-start/line-end rules. Provider `whitespace` hints cannot create a break, and
Japanese kana/kinsoku, an unmodeled class, or an unbreakable sequence becomes `text.wrapUnavailable` plus a
visible text-body placeholder. A single internal U+0020 is consumed, rather than
painted at either line edge, only when that exact soft-wrap opportunity is used;
`consumedSoftSeparators` retains its authored run and UTF-16 range.
Leading, trailing, consecutive/preserved spaces and U+3000 remain unchanged when
the line fits; a required ambiguous break refuses instead of inventing spacing.
No-wrap retains one exact shaped line even when it overflows and bypasses cluster
atomization entirely.

Each native line currently requires identical shaped ascent, descent, and line gap
across its fragments. Mixed metrics become `text.metricsUnavailable` until an
Office-qualified leading/baseline aggregation rule is available. Top anchoring is
projected directly; center and bottom anchors become
`text.verticalAnchorUnavailable` for the same reason rather than claiming an
unqualified text-block height. All sums remain integer- and coordinate-bounded.
Such refused native bodies carry `fidelity: nativeUnavailable`; only a successfully
qualified body is labeled `fidelity: native`.

The native subset requires explicit alignment, level zero, and no bullet.
Self-contained runs with a resolved typeface and size layout even when the
owning slide still carries preserve-only layout/master/theme diagnostics.
Runs that still need those unresolved fonts, or that still carry a `+mj-` /
`+mn-` token, become `text.inheritanceUnavailable` instead of substituting
host defaults. Any provider failure is paragraph/body atomic: no line
containing only the surviving sibling runs is emitted. Wrapped run fragments
retain source-run `startUtf16`/`endUtf16`; their cluster offsets are relative
to the fragment text.

The text body owns the projected inner rectangle and overflow behavior. A native
text element or text-bearing shape therefore has no enclosing element clip that
could silently turn `overflow` into clipping; the slide boundary remains the final
paint clip. Legacy v1 values without additive `textBody` metadata retain the old
clipped compatibility preview and always emit `text.layoutMetadataUnavailable`.
`text.overflow` is reserved for those legacy bounds and explicitly says the preview
does not claim PowerPoint-faithful reflow.

V1 still has no script/language/direction fields, so the host injects them. Native
square wrapping currently accepts an exact horizontal LTR paragraph only. Mixed
run directions or an RTL square-wrap request become a stable
`text.wrapUnavailable` refusal rather than guessed bidi line breaking. Native
no-wrap may retain one consistently horizontal LTR or RTL shaped line.

For legacy bodies, when runs in one paragraph request different injected
directions, the stable `text.bidiUnavailable` warning records that the tree cannot
faithfully reorder the paragraph. The preview retains native run order instead of
silently applying a browser bidi policy.

The shaping boundary accepts complete ascending clusters for LTR and complete
descending clusters for RTL, with glyph ranges kept in provider paint order.
Glyphless, zero-advance clusters remain representable for intentionally invisible
content. `ttb` and `btt` are valid shared text-contract directions, but RenderTree
v2 has no vertical line model; those runs emit `text.verticalUnsupported` and a
refusal placeholder instead of being laid out horizontally.

## Native table boundary

RenderTree v2 distinguishes legacy single-paragraph table previews from exact
cell text bodies. An exact native table requires a complete row-height vector,
column and row track sums equal to the frame extents, and one consistent text
authority mode for every cell. Its explicit cell margins define the local text
rectangle. All cell fills paint before any glyph command, so text with modeled
horizontal overflow is not covered by a later neighboring background. Exact
cells and their table add no element/cell clip; the slide clip remains. Legacy
cells keep their individual bounded compatibility clip. Tables nested in groups
inherit the same exact affine and source order as other native children.

## Website migration

Migration is intentionally one-way:

1. The package-import path loads a validated `NativePptxDeck`; legacy `WireDeck`
   values are not cast into native v1.
2. A worker or server host injects licensed font resolution and shaping, then
   compiles the visible slide.
3. The website replaces both `DeckView` and `DeckCanvasView` with one component
   that replays RenderTree paint commands into its host-owned surface.
4. Hit testing and selection use `sourceElementId`, bounds, transforms, clips, and
   z-order from the same tree rather than a parallel layout model.
5. Compatibility diagnostics and placeholders are shown in the editing UI before
   enabling operations that could imply unsupported fidelity.
6. Once gateway and compatibility fixtures use native v1, the old views and
   `WireDeck` path can be removed in a separate migration PR.

Until that migration completes, `packages/slides/src/DeckView.tsx` and
`DeckCanvasView.tsx` are feature-frozen. They may receive security or critical bug
fixes, but no new Office-fidelity work. New shape, text, table, chart-preview, and
asset behavior belongs in the native contract or this RenderTree core so the DOM
and canvas implementations cannot diverge again.

## Resource and determinism rules

Compilation validates the full native deck before invoking a provider. It bounds
tree depth, all semantic nodes (including cells, paragraphs, and runs), glyphs,
coordinates, referenced asset bytes, preset paths, and recorded commands. Provider
outputs are range-checked before conversion. Square wrapping streams compact
run/cluster ranges; it does not allocate one atom or substring per cluster, and
fragment glyph/cluster cursors advance monotonically so work is proportional to
validated input plus emitted RenderTree output.

The tree is deeply frozen. Ordered arrays preserve document order; referenced
assets sort by durable ID; canonical JSON sorts map-like object keys. The fixture
suite pins a canonical SHA-256 and checks LTR/RTL rich text, exact native body
insets/defaults, cluster-safe wrap/no-wrap, vertical anchors, overflow/refusal,
groups, every supported preset path, native solid fills/outlines,
connectors/arrows, pictures, opaque chart previews, tables/borders, z-order,
clipping, refusal paths, bounds, and the absence of UI/browser runtime dependencies.
