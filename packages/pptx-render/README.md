# @injoffice/pptx-render

Deterministic, renderer-neutral PPTX preview compilation for InjOffice native
decks. The package accepts only validated `NativePptxDeck` values from
`@injoffice/pptx-native` and compiles one slide into an immutable
`pptx-render-tree/v2` tree using integer EMU geometry. V2 adds authoritative
multi-paragraph table-cell text bodies; hosts must branch on the version before
reading table-cell text fields.

The runtime has no UI framework, markup layout engine, browser global, device
pixel input, or drawing-surface allocation. It is safe to run in a worker or on a
server. Hosts may record paint commands or replay them through the generic
Canvas2D command adapter using a context they created themselves.

Pictures can retain `clip: 'roundRect'` for the exact default DrawingML preset
(an empty `avLst`). Its radius is `min(width, height) * 16667 / 100000`, rounded
once to integer EMU. The new `clipRoundRect` paint command clips the local shape
frame after source cropping and composes with parent transforms. Host adapters
must implement this command; unsupported presets and nonempty adjustment lists
remain visible placeholders rather than unmasked images.

Parsed default pentagons use the DrawingML preset guide equations, rounded once
to integer EMU, rather than an inscribed regular polygon. Supported solid theme
fill/outline references are resolved by the native extractor. These source-bound
projections are read-only. Unsupported shape text is explicitly omitted while
independently supported geometry remains visible; diagnostics must be shown by
the host. Vertical text flow, autofit, and pentagon text-region placement are not
qualified by this geometry support. This is a partial preview, not a claim of
complete slide or Microsoft Office fidelity.

```ts
import {
  compileNativePptxSlide,
  createRecordingPaintSurface,
  paintSlideRenderTree,
} from '@injoffice/pptx-render'
import type { NativePptxTextLayout } from '@injoffice/pptx-render'
import type { NativePptxDeck } from '@injoffice/pptx-native'

declare const deck: NativePptxDeck
declare const textLayout: NativePptxTextLayout

const tree = await compileNativePptxSlide(deck, 'slide-id', { textLayout })
const recording = createRecordingPaintSurface()
paintSlideRenderTree(tree, recording)
const commands = recording.finish()
```

## Text boundary

`NativePptxTextLayout` requires the versioned font manifest, font resolver, and
text shaper from `@injoffice/font-metrics/layout`. Every run is resolved, loaded,
and shaped through that boundary. Glyph placement converts integer milli-points
to integer EMU. The core never estimates width from character count or delegates
measurement to a browser.

The native run's optional `language` retains a bounded authored language tag and
is passed to shaping ahead of host language defaults. An explicit host
`resolveRun` language override still takes precedence. This does not qualify
presentation-default cascade rules, autofit, vertical flow, or bullet font
substitution; those remain separate rendering boundaries.

Resolver, load, and shaper results are untrusted runtime inputs. The compiler
normalizes exact known fields into fresh bounded objects, validates face identity,
digests, decisions, attempted face IDs, metrics, glyphs, and clusters, and refuses
unknown or malformed provider values before they can enter RenderTree.
It snapshots the validated deck plus manifest/provider identities before the first
call, requires an explicit manifest digest, hashes loaded bytes against the resolved face, derives line metrics from those exact bytes'
design metrics, and caches one isolated authoritative plus one provider-facing
resource per exact face. Resolve/load/shape calls, unique resources, individual
and cumulative font bytes are all bounded (`64 MiB` per resource, `128 MiB`
combined authoritative/provider-facing bytes, and at most `256` unique loads).

PPTX native v1 carries exact integer text-body insets, square/no horizontal wrap,
top/center/bottom vertical-anchor metadata, fixed no-autofit, and fixed overflow. The
compiler projects that inner rectangle and wraps only at complete safe UTF-16
clusters returned by the shaper and classified as allowed from their actual text
by `@injoffice/font-metrics/layout`. The conservative v1 classifier handles
NBSP/glue, word joiner, ZWSP, hyphen/slash, Han/Hangul boundaries, and common
East-Asian opening/closing punctuation; Japanese kana/kinsoku and other
unmodeled classes refuse. It ignores provider whitespace guesses and refuses an unmodeled
line-break class. It never measures browser text or splits a surrogate pair,
ligature, or unsafe cluster. An overfull unbreakable sequence is
`text.wrapUnavailable` and paints as a visible placeholder. A single internal
U+0020 is consumed only when its exact soft-wrap opportunity is used, so it cannot
paint or skew center/right alignment at a line edge; `consumedSoftSeparators`
retains its authored run and UTF-16 range. Ambiguous leading, trailing,
consecutive/preserved spacing and U+3000 stay unchanged when they fit and refuse if
a guessed break would be required. No-wrap remains one exact shaped line, may
overflow, and bypasses cluster wrapping structures.

Native line boxes require identical shaped ascent, descent, and line gap across all
fragments. Mixed metrics visibly refuse as `text.metricsUnavailable`; center/bottom
anchoring likewise refuses as `text.verticalAnchorUnavailable` until an
Office-qualified leading and text-block-height rule exists. Top anchoring uses the
validated natural metric box directly.
Hosts may explicitly select `lineLayoutPolicy: 'max-run-natural-v1'` on
`compileNativePptxSlide`. This measured, deterministic policy takes each line's
maximum ascent, minimum descent, and maximum line gap from digest-bound shaped
fragments, shares the resulting baseline across those fragments, and sums line
boxes for vertical anchoring. Center offsets floor half-EMU remainders; bottom
offsets use the full remainder, including negative offsets for overflowing text.
Outputs carry `fidelity: 'deterministicNative'`, the policy identifier, and an
explicit warning: these rules do not establish Office pixel equivalence.
Omitting the option preserves the existing strict refusals. Unknown policies,
unresolved fonts, unsupported bidi/wrapping, and malformed provider metrics still
refuse; this option does not authorize font substitution or file mutation.
Refused native bodies report `fidelity: nativeUnavailable`; only successfully
qualified layout reports `fidelity: native`.

Native bodies also require explicit alignment, list level zero, no bullet, and
zero/absent paragraph margin and indent. The extractor now retains authored
marker/indent metadata and resolves local list/paragraph/run style defaults,
but marker shaping and indentation are not promoted to exact layout by that
data extraction alone. The separately labeled approximate file preview uses
the retained marker and paragraph offsets.
The explicit `max-run-natural-v1` policy also supports source-defined margins,
positive first-line indents and hanging character bullets. Markers are shaped
from the first source run's exact face and rendered once, at `margin + indent`,
on the first line's measured baseline. Content starts at `margin`; continuation
lines retain that margin. Non-list first lines add their source indent to the
text origin. Wrapping uses each line's resulting available width. Marker runs
carry `sourceRole: 'paragraphBullet'`; their UTF-16 offsets refer to the authored
marker, never to the content run. Explicit offsets are required for nonzero
list levels. RTL/centered bullets, absent markers, or an indent too small for the
measured marker still refuse. No numbering, tab stop, or indentation is invented.
Self-contained resolved typeface/size runs layout while leftover
layout/master/theme diagnostics stay preserve-only; missing fonts or unresolved
`+mj-`/`+mn-` tokens become `text.inheritanceUnavailable` rather than host
defaults. Those semantics refuse before any provider call instead of painting an
unqualified paragraph. A provider failure in any run refuses the complete body;
no successfully shaped sibling run is retained as a partial paragraph. Cluster
offsets are fragment-relative while each fragment's `startUtf16`/`endUtf16` map
it back to the authored source run.

A native text body owns its inner rectangle and overflow semantics, so its
enclosing text element/shape does not add a contradictory clip. Older additive-v1
values without `textBody` remain valid but emit `text.layoutMetadataUnavailable`
and use the explicitly non-faithful clipped compatibility preview.

Exact extracted table cells use the same shaped native text-body boundary. Cell
paragraphs and body metadata are an inseparable authority pair, while `text` is
only a validator-bound newline-joined summary. Their explicit DrawingML margins
define the inner text rectangle; exact overflow is not replaced by a fabricated
cell clip. Legacy authored table cells without that pair keep their bounded,
clipped compatibility preview. Table backgrounds and glyphs compile to ordinary
renderer-neutral paint commands; neither table geometry nor text uses HTML or
DOM layout.

V1 still does not carry script, language, or direction; a host injects them in
`resolveRun`. Native square wrapping currently requires one horizontal LTR
direction. Mixed run directions or RTL square wrapping refuse native line layout
instead of inventing paragraph bidi metadata. Legacy bodies without layout metadata
retain the `text.bidiUnavailable` compatibility warning.
Horizontal RTL shaping accepts the text shaper's complete descending cluster
order while retaining returned glyph paint order. Vertical `ttb`/`btt` runs are
explicitly refused as `text.verticalUnsupported` and paint as placeholders until
the tree models vertical line progression.

## Fail-closed rendering

Pictures retain exact positive DrawingML source crop insets as optional `crop`
on image nodes and image paint commands. Each inset uses 1/1000 percent
(`100000` is the full source dimension); opposing sums are strictly below
`100000`. Hosts must use the source rectangle
`(width*left/100000, height*top/100000,
width*(100000-left-right)/100000, height*(100000-top-bottom)/100000)`
when drawing the original decoded image into the destination rectangle. Do not
round to source pixels before sampling or ignore `crop`. Asset bytes, digests,
destination geometry, transforms and z-order remain unchanged. Negative/outset
and degenerate crops remain preserve-only and paint as visible placeholders.
This additive v2 field requires crop-aware image adapters for cropped inputs;
uncropped image commands are unchanged. The playground's native-file SVG preview
uses the same normalized source viewport.

- Refused elements become visible placeholder nodes.
- Refused AutoShapes need no fabricated preset: element-scoped refusals compile
  to durable-ID placeholders, while an independent slide-level refusal still
  blocks the whole slide.
- Preserve-only content always emits a diagnostic; opaque charts paint their
  exact preview image (bytes + EMU) when present and refuse with a placeholder
  when absent rather than inventing a chart renderer.
- Text shaping refusal atomically refuses the body and becomes a recorded
  placeholder command; partial sibling runs are never painted.
- Asset references retain stable IDs, content digests, byte lengths, and a
  truthful `resolutionSource`. Paint adapters resolve `assetId` plus digest from
  the validated source deck when it says `sourceDeck`, or from a host store when
  it says `host`; RenderTree never claims to contain the bytes. Host-backed assets
  also produce a resolution diagnostic. Before decode or paint, the adapter must
  verify resolved bytes against both `sha256` and `byteLength`.
- Node, depth, glyph, paint-command, provider-call, unique-font-byte, coordinate, path, and referenced-asset
  budgets prevent an accepted package from expanding without bound.

Exact native shape outlines retain width, cap, join, solid dash, and miter limit
in RenderTree and recording paint commands. Paint adapters must enact those
properties; the core does not delegate line semantics to browser defaults. Outline
and table-border widths are checked against the host coordinate budget before they
enter the immutable tree, even after native-contract validation.

Object arrays retain native z-order. Asset manifests are sorted by durable ID,
map-like object keys are sorted by `stringifySlideRenderTree`, and canonical
snapshots make drift visible in tests.

See [PPTX-RENDER-TREE.md](../../docs/PPTX-RENDER-TREE.md) for the website migration
boundary.

`sourceFrameAutoFitPreview: true` separately opts into rendering explicitly
marked `shape-source-frame` text bodies. Without it, those bodies remain refused
even when `lineLayoutPolicy` is set. Opted-in bodies report
`fidelity: 'approximateSourceFrame'` and a warning; they use the original frame
without resizing and do not qualify Office-equivalent layout or editing rights.
