# @injoffice/docs

### Read-only partial source-content projection

The optional same-byte DOCX inspection response also carries source-bound basic
equation evidence. `createNativeDocxEquationPreviewsV1` validates its package,
paragraph, original unsupported diagnostic and exact anchor before exposing a
bounded JSON tree. Rows, literal math text, fractions, superscripts, subscripts
and explicit square roots can be shown by a host through fixed MathML elements.
Unknown formatting, revisions, attributes, constructs, malformed child order or
unqualified visibility omit the whole equation. This is browser math layout,
not Word typography, font fidelity or native pagination. The demo lists these
equations separately under its explicit read-only partial-text preview. Strict
painting and equation-paragraph mutations remain refused. No HTML strings,
links, arbitrary MathML attributes or source font assets are exposed.

An optional `equation_context_notices` inventory on same-byte inspection can
qualify exact horizontal `lrTb` section direction, explicitly suppressed
paragraph hyphenation, bounded non-drawing tab stops and two source-declared
charset pairs for this equation-only view. These diagnostics remain in the
original document/layout; neither ordinary text recovery nor native paint uses
the exception. Charset declarations are ignored font-matching metadata, not a
reinterpreted encoding or a claim of consistent/default font selection: MathML
uses the already XML-decoded Unicode. The demo displays the retained notice count.
Standalone hosts must supply trusted same-byte extractor evidence. The decoder
joins styles/font-table digests to passthrough part hashes and bounds their slice
anchors; it does not recompute XML slice hashes without source bytes. Main-part
whole-part SHA is producer evidence; its original section diagnostic anchor
(including slice SHA), section owner and package identity are joined exactly.

`createNativeDocxPartialContentPreviewV1(document,
{ policy: 'source-text-with-omissions-v1', read_only: true }, resolvedLayout?)`
is available from the browser-safe `@injoffice/docs/native-docx` entry. It
validates the native source model and, when supplied, the resolved model's exact
document/revision/main-part identity. The result is a continuous plain-text
source projection with source-bound omission placeholders, **not paginated
native paint, an editable document, or a substitute persistence format**.
Text segments are literal source strings: render them as text, never as HTML.

The profile includes ordinary body text and unmerged table-cell source text,
grouped by source row/cell ordinals without reconstructing table geometry.
Qualified authored drawing descriptions appear as labeled alternative text
alongside the retained drawing omission. Hidden text, fields, controls, drawing
geometry, merged cells, table layout and nonbody stories receive placeholders;
source diagnostics remain attached. Unknown document-wide diagnostics prevent
text qualification. Without resolved layout, inherited visibility is unknown,
so the result provides an inventory only and exposes no text. Formatting,
numbering and document layout are not reconstructed. A 200-body-block and
200-table-cell and 100,000-text-unit budget produces explicit truncation placeholders rather than
silently dropping source content. The original source and mutation guards are
unchanged; package hashes remain trusted native-extractor evidence, not XML
bytes independently re-hashed by this consumer.
The resolved-layout V1 contract has no package digest of its own: the caller
must supply document and layout from the same authoritative extraction. The
identity join is not independent proof that arbitrary caller-provided models
were extracted from particular package bytes.

The playground displays this inventory separately from its existing richer
continuous editor; limitations of this reusable projection do not imply that
the editor omits the same content.
Its explicit partial-text button uses `@injoffice/docx-wasm` inspection to
resolve the same source bytes locally, even while the editor uses server mode.
It never uploads source bytes or changes editing authority. Equations remain
unsupported: no cached equation text is invented from unmodeled XML.

Native DOCX text shaping honors inherited `w:kern` minimum sizes in bounded whole half-points (1–3276). The resolved `kerning_min_size_half_points` threshold enables kerning when the resolved `w:sz` is at least that threshold; absent thresholds explicitly disable kerning, as specified by ECMA-376 §17.3.2.19. Direct kerning markup remains source-preserved and does not grant replacement permission. Unit-bearing, malformed, duplicate, and out-of-range thresholds remain unqualified. This may change advances from earlier previews that inherited HarfBuzz's default kerning without an authored DOCX setting.

Renderer-independent Docs contracts and layout helpers. Native DOCX v1 is the
strict JSON boundary between authoritative `go/docxpatch` parsing/persistence
and future browser layout and ProseMirror adapters.

```ts
import { decodeNativeDocxJson, encodeNativeDocxDocument } from '@injoffice/docs'

const decoded = decodeNativeDocxJson(await response.text())
if (!decoded.ok) {
  console.error(decoded.issues)
} else {
  const canonical = encodeNativeDocxDocument(decoded.value)
}
```

Browser-local extract/apply clients should import the native contract and
mutation envelope from `@injoffice/docs/native-docx`. That entry does not load
the root package's Node-qualified shaping or page-paint providers.

The published schema is available as
`@injoffice/docs/native-docx-v1.schema.json`. Unknown fields, dangling native
identities, wrong-kind references, inconsistent source anchors, invalid union
payloads, unsafe edit policies, JSON null/negative-zero hazards, and inputs over
the documented resource limits are rejected.
The model is not a DOCX generator: source anchors and passthrough inventories
must be applied surgically to the original OPC package by `go/docxpatch`.

## Native ProseMirror transaction adapter v1

`adaptNativeDocxProseMirrorTransactionV1` translates only a complete, exact,
versioned body-text projection into the existing shared Office mutation
envelope and `go/docxpatch` run-replacement payload. Positions are ProseMirror
UTF-16 code units. The input binds the exact `source.package_sha256`, durable
paragraph/run IDs, run XML fingerprints, native run properties, opaque mark
sets, every pre-step run value, and the complete before/after body projection.

V1 accepts only closed text replace slices wholly inside one native text run.
It simulates successive step coordinates, rejects mapping drift, overlap and
reordering, and coalesces ordered changes to one guarded mutation per run.
Tables, controls, references, drawings, notes/comments, fields, tracked or
unsupported markup, mark changes, structural/open slices, stale revisions or
anchors, unknown or duplicate keys, unsafe XML text, unattested edge
whitespace, partial coverage, and all resource
overflow return issues with no envelope. The canonical output uses the shared
3 MiB payload and 4 MiB envelope ceilings.

The typed `NativeDocxTextMutationPayloadV1` also exposes Go's narrow direct
paragraph selector. `target_kind: 'paragraph'` is valid only when the native
paragraph contains exactly one run and that run is text; it binds the
paragraph ID and paragraph XML fingerprint while Go replaces the sole text
node. It must not overlap that run's selector in the same atomic batch. The
ProseMirror adapter continues to emit run selectors because its steps and
source fingerprints are run-addressed.

The optional `adaptNativeDocxProseMirrorTransactionWithHostV1` seam accepts a
host-owned projector for an actual ProseMirror transaction. It adds no runtime
ProseMirror dependency and no persistence or layout authority: the original
DOCX and Go native engine remain authoritative. The website/editor integration
must carry durable native IDs and exact anchors in its TipTap schema, project
transactions before save, and keep collaboration sequence/rebase state
separate from the exact package-byte revision.

## Native shaping and lines

The additive `shapeNativeDocxLinesV1` API consumes both validated Go wire
projections (`NativeDocumentV1` and `NativeResolvedLayoutInputV1`) plus the
shared `@injoffice/font-metrics/layout` manifest/provider contract. A future
paginator supplies an explicit inline width and tab interval; the core never
reads browser pixels or invents Word defaults.

```ts
import { shapeNativeDocxLinesV1 } from '@injoffice/docs'

const result = await shapeNativeDocxLinesV1(
  {
    protocol: 'injoffice.docx.shaping-request',
    version: 1,
    document: nativeDocument,
    resolved_layout: resolvedLayout,
    font_manifest: manifest,
    available_width_millipoints: 468_000,
    tab_interval_millipoints: 36_000,
  },
  { resolver, shaper },
)
```

The output is `injoffice.docx.shaped-lines` v1: deterministic paragraphs,
lines, and cluster fragments keyed to native paragraph/run IDs. All distances
are integer milli-points and text offsets are UTF-16. The core handles resolved
paragraph spacing/indents/alignment/line rules, paragraph-content-relative
tabs, hard breaks, exact resolved list-marker text and geometry, script/language/direction
shaping requests, and conservative cluster-safe Unicode wrapping.
The Go resolver discovers `numbering.xml` only through the main-part internal
relationship and exact content type, then attests the owning relationship part,
raw numbering part, and canonical source-ordered marker model. The bounded
slice supports decimal, lower/upper repeated-letter, lower/upper Roman, and
literal bullet levels; `%1` through `%9` placeholders with other percent forms retained literally,
concrete-instance `startOverride`, bounded restart policies, authored numbering suffix-tab stops,
logical alignment at the numbering text-margin anchor, and hanging indents are
resolved before HarfBuzz shaping. Style-supplied `ilvl` is ignored in favor of
one exact abstract-level `pStyle` link, and definition/model hashes use the same
versioned canonical UTF-8 serializer in Go and TypeScript. Explicit numbering tabs
at or before the shaped marker end and marker advances beyond the label end are
refused before any partial marker can escape. Pagination recomputes the shaped
marker end and suffix target from the exact numbering tab or attested Word default
tab interval, so a self-consistent shaped-marker tamper cannot pass the join.
Unknown or foreign numbering-root semantics stay silent only while numbering is
unused and block every concrete `numId` reference. All accepted twips are bounded
before their exact ×50 milli-point conversion, with negative zero rejected.
Resolved paragraph-mark properties supply font metrics for blank,
hidden-only, tab-only, and explicit empty lines without emitting synthetic
text fragments.

The bounded visual slice resolves Unicode bidi over the logical paragraph with
an explicit LTR/RTL paragraph base, applies explicit run direction as UAX #9
isolates, shapes each resolved direction/script span, and emits fragments in
left-to-right visual paint order. Every fragment retains its bidi level and
logical ordinal; each line carries the complete logical-to-visual permutation.
Soft-wrapped `both` lines expand only shaped U+0020 clusters between content,
distributing integer milli-point remainders in visual order and recording each
expansion. Final and hard-break lines remain unexpanded. `distribute`, authored
bidi controls, and soft justified lines without a qualified U+0020 opportunity
fail closed rather than using character spacing.

Table grids, drawings, theme/script gaps, pagination, and painting remain
explicit diagnostics. Both
wire projections and the font manifest are strictly validated before a
provider runs; injected font resources and returned cluster/glyph ranges are
validated again, including manifest-backed face identity, bounded metadata,
and the loaded byte SHA-256. A refusal or invalid result in any numbered
paragraph atomically empties the complete shaped paragraph projection, so a
partial list is never presented as complete. The validated request/manifest and provider callable
identities are snapshotted before any provider call; output records those exact
manifest and resolver/shaper IDs and revisions, and mid-call provider identity
changes fail closed. Validated resources are cached by the complete resolved-face
identity: at most 64 MiB per resource, 128 MiB of unique font bytes, and 256
faces, with 50,000 resolve/shape calls and 256 loads per request. Resolver and
shaper results, runs, and faces are copied into owned snapshots at each async
boundary. The shaper receives one separately copied resource per exact face,
bounded to another 128 MiB per request, so its mutable `Uint8Array` cannot
corrupt the digest-checked authoritative cache. Because JavaScript cannot make
typed-array elements immutable without copying on every span, the injected
shaper remains trusted not to change that provider-facing byte copy; all
identity, metrics, run, manifest, and provenance metadata stays immutable.
The module imports no DOM, HTML, canvas, Mammoth, React, Konva, or platform
font-measurement API.

For Node 22, `@injoffice/font-metrics/harfbuzz` supplies the resolver-neutral,
real HarfBuzz qualification boundary for page-paint v1's horizontal
Latin/Arabic/Hebrew slice.
It consumes only the exact face/bytes/metrics returned by the injected
resolver; it performs no inventory, resolution, loading, or substitution.
Its composite `providerRevision` binds the pinned HarfBuzz JavaScript/WASM
artifacts, runtime, generated Unicode 13 classifier, configuration, and host source revision.
The page-paint compiler accepts only a shaper instance attested by that exact
runtime module and resolve/loads every unique authored font reference through
the digest-bound embedded-font provider before bidi resolution begins.
The shaping cache key binds the exact digest and face; shaped-lines carries the
face ID plus manifest/provider revisions; page paint's full manifest and
shaped-lines hashes complete that content-addressed join. Shaped-lines also
records the pinned bidi provider revision, Unicode 13.0.0 data version, and
hash-bound shared classifier revision. The bidi boundary atomically refuses all
supplementary scalars before invoking the UTF-16-based engine.

## Native pagination v1

`paginateNativeDocxV1` is the next renderer-neutral stage. It consumes the
native document, resolved-layout input, and shaped-lines output together, then
emits `injoffice.docx.paginated-layout` v1. The request and output have strict
exact-key decoders; document/revision/story/paragraph/run/fragment joins are
verified before placement.

```ts
import { paginateNativeDocxV1 } from '@injoffice/docs'

const result = paginateNativeDocxV1({
  protocol: 'injoffice.docx.pagination-request',
  version: 1,
  document: nativeDocument,
  resolved_layout: resolvedLayout,
  shaped_lines: shapedLines,
  pagination_settings: paginationSettings,
})
```

The supported v1 slice places shaped body paragraphs and lines in integer
milli-points (one twip is exactly 50 milli-points). It applies page geometry,
top/right/bottom/left margins and gutter, continuous/next/even/odd/next-column
section starts, exact equal-width column grids,
paragraph page breaks, keep-lines, atomic keep-next chains, default-on widow/orphan
control, and Word's maximum (not additive) adjacent paragraph spacing. Stable
page, paragraph-slice, and placed-line IDs retain section and shaped-line
provenance. Blank parity fillers are explicit, header/footer-free output pages
owned by the preceding section and point at the odd/even section they precede.
Consumers that receive stored or transported page output should use
`decodeNativeDocxPaginatedLayoutForRequest(output, request)`: unlike the
standalone structural decoder, it proves every shaped body line is present
exactly once in native source order and binds source sections, page geometry,
header/footer provenance, engine provenance, vertical placement, paragraph
spacing, and page/slice break choices to the exact deterministic request. The
paginator runs the same source-completeness check before returning success.

An exact bounded note slice is included. Footnote and endnote stories must be
relationship-resolved from their owning main-part relationship, carry distinct
native IDs, and contain paragraphs only. Each content story must have exactly
one matching native reference in the body and exactly one self-label reference;
decimal labels are assigned from body reference order, independently per note
kind, starting at one. One qualified separator story is required per used kind.
Footnotes are placed as an atomic separator-plus-notes group in the unused
bottom region of the referencing page. Endnotes are placed as one atomic group
on the final content page, or on one new final page when the existing page has
insufficient room. Every placed note and note line carries its exact section
and column identity. The normal `-1` sentinel produces one deterministic
bounded separator rule; an unused `0` continuation sentinel stays inert.
Actual continuation still refuses the whole projection.

Because keep-with-next constrains a paragraph boundary rather than making every
line indivisible, v1 accepts a keep chain only when every multiline member is
also `keep_lines`; other chains refuse pending boundary-aware split planning.

V1 is deliberately fail-closed. Every section column must have the same body
width used by shaping. Continuous transitions require an identical
single-column physical grid; next-column requires an identical exact grid.
Equal-width columns must divide exactly, while explicit columns must completely
describe equal widths and authored gaps; omitted explicit spacing is standard
zero. Notes that require body reflow or splitting refuse the complete
projection; an unused continuation separator remains inert. Custom
numbering/restarts/positions, ambiguous or duplicate
references, missing labels or separators, cycles, nested tables, drawings,
fields, unsupported note markup, note-bearing pages with multiple columns or
section grids, and cross-continuous note boundaries also refuse with no partial
pages. Note-free multicolumn documents remain supported. Column separators and
unequal or ambiguous widths,
unsupported attested settings semantics, body tables, inline/floating drawings,
and body comment or non-note references,
inline page/column breaks, and invalid/unknown pagination-affecting source
markup return `status: 'refused'` with no partial pages. Pagination retains
header/footer reference provenance; the native header/footer planner resolves
per-kind section inheritance and selects first/even/default stories from exact
`title_page`, physical parity, and the attested even/odd setting. The separate extractor-owned settings attestation represents an absent
part with exact Word defaults, but pagination refuses it because omitted
compatibility mode means Word mode 12. The bounded supported profile requires
explicit mode 15, binds the source package plus the owning `.rels` and settings
part fingerprints, and verifies that shaping used its default tab stop. Mirror
margins, top gutter, legacy/unknown compatibility semantics, and unknown
layout-affecting settings refuse. Every accepted settings element has an exact
attribute/child/text shape; theme-font language, shape defaults, attached
templates, native math defaults, placeholder/revision display settings,
enabled field-result updates, and other inputs not proven irrelevant to
resolved/shaped advances remain unsupported. `updateFields=false` is attested
explicitly, and content-type identity uses ASCII-only case folding. Explicit
extractor-owned Word default section
geometry is accepted; no other geometry is inferred.

### Explicit operator font substitution (read-only)

`renderNativeDocxFontSubstitutionPreviewV1(input, outlineProvider, { fonts })`
is a separate approximate compiler. Supply a content-addressed host `manifest`,
`resolver`, and `substitutionPolicy` on `fonts`; the policy is
`{ version: 1, mappings: [{ sourceFamily, targetFamily, weight: 400 | 700,
style: 'normal' | 'italic' }] }`. Exact supplied faces always win. There are no
built-in substitutions, system font discovery, synthetic styles, or bundled
proprietary fonts. `compileNativeDocxFontSubstitutionPreviewV1` accepts the
corresponding original-source shaping/page-paint request plus the same policy.
Neither function exports a strict prepared artifact or grants editing rights.

The distinct `injoffice.docx.font-substitution-preview` envelope carries the
original source identity, policy/hash, actual source-to-selected font records,
the selected font manifest joined to rendering provenance, and persistent
approximation warnings. `decodeNativeDocxFontSubstitutionPreviewV1` validates
this output (also exported from the browser-safe `native-page-paint-output`
entrypoint). Source font names, package bytes, and unrelated diagnostics remain
unchanged; removing or forging substitution evidence cannot qualify strict paint.

This first tier requires modern Word settings and graphic ASCII/LTR source runs
or Latin numbering, plus eligible blank paragraph-mark metric consumers. Symbol
bullets, font-matching metadata, EA/CS/RTL selection, legacy layout, automatic
borders, absent-size policies, page fields and square-wrap combinations remain
unsupported. Layout is not Word-validated; an explicit request may still return
`status: 'refused'` with no pages. These constraints currently prevent applying
this font-only tier to the original benchmark documents with matching metadata
and legacy settings.

The separate opt-in `renderNativeDocxApproximatePagePreviewV1` path may use
current layout rules for exact legacy mode 12/14 settings. Its extractor-owned
eligibility can additionally retain bounded known theme-language, locale,
math-default, shape-ID and compatibility-flag facts in `approximated_settings`.
These are explicitly disregarded settings, not implemented Word semantics:
each carries a warning, its original values and source path, while original
strict diagnostics and source hashes remain attached. Unknown or malformed
settings still refuse. Active unsupported math, VML content, missing fonts and
unsupported geometry are not made renderable by this settings policy. The
result remains a distinct, read-only approximate envelope; strict pagination
and mutation safety are unchanged.

Current-layout approximation reserves expanded line boxes using natural ascent
from the top and leaves extra leading below the text. This declared host policy
is not Word baseline fidelity; strict rendering still requires natural line
height, and compressed line boxes remain refused in both paths.

Eligible mode-12 previews can also carry `legacy_table_origins` source facts.
For a qualified unmerged, left-aligned table in a single page column, the
read-only `legacy-content-aligned-origin-v1` policy shifts the table and its
cell content left by its explicitly authored leading cell margin. It does not
change widths, row heights, source indents, or document bytes. The source
indent/margin paths and part digests enter the table projection hash; placement
is re-derived and bounded against the page. Mode 14/15, ambiguous source
properties, and page-underflow cases do not receive this policy. The output
retains the facts and a visible approximation warning, including when combined
with automatic-border preview. This is not a Word-layout equivalence claim.
Modeled main-part digests are trusted native-extractor evidence joined to the
package; the TypeScript consumer does not independently hash absent XML bytes.

Mode-12 current-layout previews additionally use the declared
`collapsed-horizontal-border-reservation-v1` policy for automatic-height,
unmerged tables with one shaped line per cell and equal explicit single top,
inside-horizontal and bottom borders. Each row reserves one authored border
width above its content, independently of font metrics and cell padding.
This bounded approximation follows collapsed-border space accounting, not a
Word-validated baseline rule. Unequal borders, multiple lines, explicit row
heights and other unqualified geometry retain their existing behavior. The
named policy enters the table projection hash and deterministic placement
replay; strict rendering and source bytes remain unchanged. The approximate
envelope declares this policy alongside its other fidelity warnings.

The same current-layout approximation can accept an explicit host-selected
`fontSizePolicy: { kind: 'host-default-size-v1', half_points: 22 }` only for
native-source-attested missing sizes (`absent_font_sizes` eligibility facts).
This is an 11 pt consumer choice, not an authored or Microsoft default.
The output retains the source omissions, chosen sizes and a visible warning;
existing sizes and malformed/unsupported source diagnostics are never replaced.
Strict rendering continues to refuse missing required font metrics.

`renderNativeDocxAutomaticBorderPreviewV1` is a separate opt-in read-only
contrast policy for native-source-qualified automatic table borders. It uses
black only on a proven white preview surface with absent or exact white
effective table/cell fills. Unknown backgrounds, conflicting borders and
unqualified style effects remain refused. The source model and strict
diagnostics are unchanged; the distinct `injoffice.docx.auto-border-preview`
envelope retains the policy, source-bound evidence and original diagnostic
identities. Consumers must render its pages on the declared opaque white
surface and display its approximation warning. This is not a claim of Word
automatic-color fidelity. Optional legacy-settings eligibility is validated
independently; this rendering policy does not grant a settings exception.

Approximate body `PAGE`/`NUMPAGES` fields use the same bounded fixed-point solver
as strict layout: field text derives from final page placement, not cached
values, with cycle detection and an eight-pass limit. Original field source,
read-only policy, settings and integrity hashes remain bound to the preview.
The same solve handles page-edge square images with or without body fields,
revalidating the final source-derived exclusion intervals. It retains the
single-column, left-aligned LTR, no-table/no-note restrictions; interior islands,
fully blocked lines and nonconvergence refuse. Repeated header/footer page
fields expand from the final page count and restarted page numbers, with the
same source-bound variant coverage, fragment budgets and stale-cache refusals
as strict rendering. Fields in notes/comments and body table cells remain
unsupported.

Top-of-page paragraph-before spacing is retained only on the first content page
of a section and suppressed on later pages, including automatic, explicit, and
keep/widow-driven page moves. Keep-chain planning uses one bounded reverse pass.
The paginator is bounded to 2,048 pages, 100,000 line placements, 100,000
paragraph slices, 1,000 diagnostics, two million decoded output values, and
integer coordinates within one trillion milli-points. It imports no DOM,
HTML, CSS layout, canvas, browser renderer, Mammoth, or legacy pixel paginator.
Wire-visible diagnostics use locale-independent UTF-16 code-unit ordering.
The provenance checks are structural joins, not cryptographic authorization;
callers must keep all four projections inside the trusted native pipeline.
Paginated-layout v1 does not carry full resolved-layout/shaped-lines content,
so independent transport still requires the complete request and request-bound
deterministic recomputation. Numbered documents additionally carry the exact
numbering relationship/part hashes and canonical marker-model digest through
pagination provenance.

## Native page paint v1

`compileNativeDocxPagePaintV1` closes the final renderer-neutral gap from the
native document/resolved-layout/shaped-lines/pagination request and its exact
paginated output to deterministic replayable page paths. The request carries
the full font manifest plus canonical manifest, shaped-lines, qualified-table,
embedded-media, and complete paginated-layout SHA-256 attestations, plus an
expected outline-provider ID/revision. Every
painted `face_id` must resolve to a non-system manifest face with an explicit
SHA-256 content digest; the injected provider is called only with the tuple
`(content_digest, collection_index, glyph_id)`, with `face_id` retained as the
manifest identity attestation.

The provider returns bounded integer design-unit move/line/quadratic/cubic/close
commands and `units_per_em`. The engine applies the shared deterministic
font-unit scaling rule, the resolved run or list-marker font size/color, shaped glyph offsets
and advances, and the paginated line origin. Font outlines use a y-up baseline
coordinate system; output paths use absolute y-down page milli-points. Each
page explicitly records its white background, full-page clip, body geometry,
every placed line (including glyphless lines), baseline, nonzero fill rule, and
left-to-right visual-order glyph path commands. A replay adapter therefore performs no
text shaping, measurement, font selection, or paint-policy inference.

V1 deliberately accepts only diagnostic-free, simple horizontal body and
selected static header/footer paragraphs with natural shaped line height, plus
qualified footnote/endnote paragraphs with natural shaped line height, plus a
bounded body-table subset:
explicit fixed dxa or percentage width and grid, left alignment/indent, four cell margins,
indivisible or line-safe natural-height rows, direct cell paragraphs, table-level single/RGB
borders, and clear RGB cell shading. Horizontal `grid_span`, vertical restart/continue
merges, `atLeast`/`exact` row heights, and simple whole-table styles that project
onto those same border/fill commands are included. Merged rows paginate as one
atomic group and shared edges are emitted once in deterministic table-fill →
line-content → border order. A contiguous leading `repeat_header` row prefix
repeats on each continuation page, with the initial prefix kept together with
the first body row. Each later page must fit the complete header prefix plus
the next indivisible row; otherwise the whole preview refuses. Repeating-header
tables currently refuse vertical merges, rather than splitting ambiguous merge
groups across the header/body boundary. Source-bound cell IDs distinguish the
parallel text flows of adjacent cells; repeated line placements have unique
page-derived IDs and are validated by exact pagination replay, not accepted as
arbitrary duplicate body content. Generated two-column DOCX browser fixtures
verify this bounded behavior, not Word pixel equivalence.
Natural-height rows without `cant_split: true` use a line-safe fragmentation
policy: cuts cannot bisect any adjacent cell's shaped line, `keep_lines` group,
or initial/final two-line widow group. Header prefixes repeat before fragments;
cell fills and side borders continue even when that cell has no remaining text.
Horizontal source borders appear only at the source row boundaries, not invented
at page cuts. Fragment source ranges are contiguous, identity-bound, and checked
by exact pagination replay. Split rows currently refuse explicit/minimum heights,
vertical merges, keep-next chains, forced paragraph page breaks, and documents
containing notes. This is deterministic bounded pagination, not a claim of Word
row-break parity or content-based table autofit.
Percentage width uses the named `fixed-grid-percent-exact-twips-v1` policy:
resolve the authored 1..5000 fiftieths-of-a-percent against the owning single
section column, then scale the authored grid proportionally. Every preferred
cell width must match its original grid span, and all resulting widths must be
integral twips; non-integral allocations refuse rather than silently round.
The policy, source grid, percentage and container identity enter the qualified
table hash. Cell text is shaped again at the resulting content widths. This is
flexible percentage sizing of a fixed grid, **not content-based autofit**.
Content-based sizing is a separate `shaped-content-minmax-v1` policy for explicit
`autofit` tables. Source-resolved table geometry can also select this path: an
explicit, complete unconditional table-style chain supplies inherited geometry,
then direct table properties override it. The separate resolved model retains
auto-width resets and per-side margin inheritance; source objects remain unchanged.
The documented defaults are autofit layout, auto preferred width, left alignment,
zero indent, zero top/bottom margins, and 115-twip left/right margins when absent
throughout the qualified chain. Explicit inherited margins override these defaults.
Unknown geometry, conditional chains, automatic border colors and unsupported
alignment remain outside this qualification; no built-in style or font is guessed.
A bounded preliminary pass uses the same attested fonts and
HarfBuzz shaper as the final render. Intrinsic minima come from complete
space-separated words; maxima come from complete hard-break-delimited lines.
Widths round upward to integer twips so text minima are never rounded down.
The table's absolute preferred width is clamped between these intrinsic bounds
and its owning section; omitted/auto preferred width uses the maximum that fits.
Remaining width is distributed in proportion to each column's min/max headroom,
with deterministic largest-remainder allocation and source-order ties.
For omitted/auto table width, a separate
`source-preferred-nonconflicting-v1` policy preserves the authored grid when every
cell in every unmerged row explicitly repeats its column's preferred width,
each column's shaped maximum content plus margins fits that width, and the
complete grid fits the owning section. Empty columns therefore retain their
consistent authored preferences. Explicit absolute table widths, missing or
conflicting cell preferences, and content requiring wrapping at those preferences
continue through the existing content policy and its refusal boundaries. This
bounded case does not establish general Word autofit fidelity. Its named policy
and source preferences enter the same independently re-derived table hash.
Authored grid/cell widths remain source preferences recorded in the policy;
they are not immutable column widths. Final wrapped glyph clusters independently
rederive the same allocation during source-bound pagination and paint validation.
This supports unmerged direct-text cells in one section column, including natural
row fragmentation and repeated headings. It refuses unsatisfied word minima,
percentage preferred widths, paragraph indents/justification, tabs, discretionary
breaks, special spacing controls, RTL paragraphs, numbered/merged cells, and
unqualified font/source diagnostics. This is genuine font-content sizing under
an explicit bounded policy, **not a claim of Microsoft's autofit algorithm**.
Both table policies refuse non-prefix repeating headers,
cell-border conflicts, nested content, numbered cells, conditional
`tblStylePr` effects, and any table-descendant resolution diagnostic. Selected story
lines must fit between the exact header/footer edge distance and the body box.
The planner independently inherits reference kinds, keeps parity fillers blank,
and content-addresses its complete selection/placement plan in page-paint
provenance. Exact left-to-right list-marker fragments are replayed from their
already-shaped glyphs and numbering provenance.

Text-run `vertAlign` subscript/superscript uses the explicit font-metric profile
described below; unsupported script metrics or source forms still refuse.

The qualified native-image slice is equally renderer-neutral and
self-contained: the compiler exact-joins each drawing's internal relationship
and preserved relationship-part digest to one preserved media part, verifies
caller-supplied bytes against that part's byte length and SHA-256, parses bounded static PNG/JPEG
dimensions, and carries canonical base64 bytes as an output resource. Inline
image commands retain the exact DrawingML EMU extent projected through the
integer-only `10/127` milli-point ratio, an explicit source crop rectangle, and an
explicit source-bound orientation transform. Prepared and completed compiler envelopes expose
canonical hashes for the complete validated request and output.

Inline images qualify as embedded static PNG or baseline JFIF JPEG pictures in `wp:inline` with zero
distances/effect extents, extent-preserving `a:xfrm` (quarter-turn rotation and
horizontal/vertical flips), bounded source crop, exact integer
milli-point geometry, and bounded bytes/pixels. Both inline and qualified floating
images refuse remote or external relationships, vectors and other
raster formats, animation, negative/extending crop, arbitrary rotation, effects, mismatched extents, and
media digest drift. JPEG support is deliberately bounded to one baseline 8-bit
grayscale/YCbCr interleaved scan with internal tables and JFIF APP0, following
[ITU-T T.871](https://www.itu.int/rec/T-REC-T.871). EXIF, ICC, Adobe transforms,
progressive/multiple scans and ambiguous color metadata remain refused. The
validator checks marker structure, not entropy decoding; the viewer decodes the
preserved bytes and the browser smoke verifies generated red/blue JPEG pixels.

Body-paragraph `wp:anchor` pictures also qualify with explicit page-relative
non-negative offsets, `wrapNone` or explicit `wrapSquare wrapText="bothSides"`, zero distances/effects, disabled `simplePos`
and `locked`, and enabled `allowOverlap`/`layoutInCell`. Their source
`behindDoc` and `relativeHeight` become `floating_layer` and `stacking_order`.
The anchor consumes no inline width or image-height line space; its actual
paginated paragraph selects the page. `paint_floating_image` replays behind all
page content or in front of it, sorted by the source stacking order. Consumers
must handle this additive command and respect global page replay order rather
than concatenating line-owned commands. At most 128 floating pictures qualify;
orders must be unique per layer, and images
must fit wholly inside the page. Header/footer/table anchors, alignment-based
positions, relative sizing and collision avoidance remain refused.
Square wrapping is bounded to unrotated page-edge rectangles leaving one text
interval in single-column body paragraphs with left/start-aligned LTR text and no
numbering, tables or notes. Source-derived intervals shape complete lines beside
the image and restore paragraph width below it. Interior islands, fully blocked
lines, tight/through wrapping and vertical displacement remain refused. Wrapping
and body page fields share one eight-pass, cycle-detecting pagination solve;
final paint validation independently replays the exact source exclusions. Neither
the original package nor its native source text is modified.
This is a source-contract implementation, not independently established Word
pixel parity.
Quarter turns require explicit unrotated DrawingML extents whose swapped bounds
exactly match `wp:extent`; missing or inconsistent extents refuse before painting.
Reflections apply in source axes before clockwise rotation. Integer SVG matrices
and original raster pixel fixtures verify all orientations without Office screenshots.
Source crops use integer one-hundred-thousandths, default omitted sides to zero,
and must retain at least 1% of the original image on both axes (at most 100×
scale). The viewer clips this source rectangle before reflecting/rotating into
the unchanged layout box; it does not rewrite or resample the original media.
Selected header/footer inline PNG/JPEG runs use the same
digest-bound asset join and `paint_inline_image` command as body pictures. V1
also emits RTL/mixed-bidi fragments, exact list-marker glyphs, and bounded
U+0020-justified lines in visual paint order, requiring no paint-time text
reversal or measurement. It refuses distributed-character justification,
underline styles outside the metric-bound subset below, header/footer tables,
shapes, references, fields outside the
page-number subset below, and unsupported note content. Native note marker
fragments must exactly equal the paginator-assigned decimal label. It also refuses
system or unaddressed faces, missing glyphs, invalid/mismatched provider output,
unclosed or overflowing paths, incomplete pages, and all resource overflows.
Any such condition returns one `status: 'refused'` output with `pages: []`; no
previously accumulated page or command escapes.

Text-run highlighting emits the additive `fill_text_highlight` paint command.
Consumers must replay these filled rectangles in command order before the line's
glyph paths; do not assume every non-image command is a glyph. Background width
comes from each shaped visual fragment's advance (including ordinary spaces),
and height from its qualified font ascent/descent. This deterministic native
metric policy is not a claim of Word-pixel parity. The compiler and strict
request decoder bind every rectangle to the source highlight color, run,
fragment, placement and geometry, and reject missing/reordered decorations.
The 16 OOXML named colors and `none` are supported for text runs; zero-advance
fragments produce no background. Highlighted tabs/other controls and list-marker
backgrounds remain refused. Edit authority is unchanged.

Underlined runs support `single`, `double`, `words`, and `none` through explicit
`stroke_text_underline` commands after the line's glyphs. Geometry uses the
content-addressed font's underline position and thickness; missing metrics
refuse instead of using browser decoration defaults. Double underlines have one
stroke-width of clear separation, and `words` omits whitespace fragments.
Source/style, fragment, font metrics, placement and command order are checked
again by the request decoder. Other underline styles and custom underline
colors remain outside this bounded profile. This is a deterministic metric
policy, not a claim of Word-pixel parity.

Simple decimal `PAGE` and `NUMPAGES` fields in ordinary header/footer paragraphs
are resolved from the final native body pagination. The Go extractor recognizes
an unlocked `w:fldSimple` or a flat paragraph-local five-run
`begin` / `instrText` / `separate` / result / `end` sequence, with an exact
`PAGE` or `NUMPAGES` instruction and one supported text result run. It discards the cached result, records
`page_field`, and keeps the paragraph read-only. The raw story-root XML anchor
digest covers the instruction and cached bytes; the package digest binds the
complete source. These are integrity joins within the trusted extraction
pipeline, not signatures from an external authority.

The compiler re-shapes complete header/footer text for each final page, including
surrounding text and authored alignment. Optional `page_field_variants` in the
paint request must exactly cover final page order; their source text is checked
against deterministic decimal substitution. Their full shaping data is included
in the shaped-lines integrity digest. Native glyphs paint `Page 1 of 2` and
`Page 2 of 2` even if the source cache says `999`. No cached value is rendered or
made editable. Low-level consumers must pass the compiler's variants through
to header/footer layout and paint, not reuse the empty source field as text.

This bounded profile allows at most 64 pages and 100,000 cumulative variant
fragments (`DOCX_PAGE_FIELD_LIMITS`). Note/comment fields, header/footer
tables, other complex `fldChar` sequences, nested fields, unsupported switches,
locked/dirty fields, non-decimal section numbering formats, and all
other field instructions remain refused. This implementation makes no Word-pixel parity
claim. See the [OOXML simple-field definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.simplefield?view=openxml-3.0.1).

Ordinary body-paragraph PAGE/NUMPAGES fields use a bounded whole-body layout
fixed point: start with decimal `1`, shape and paginate, derive field text from
the actual page carrying its glyphs, and repeat until the complete state is
unchanged. Cycles and failure to converge within eight passes refuse. At most
128 body fields and 64 pages qualify. Across solver passes and subsequent
header/footer variants, at most 100,000 shaping fragments are processed. Hidden fields, table/note/comment fields
and fields split across pages refuse. No stale cached result seeds the solve.
The paint request retains `body_field_source` plus its integrity digest and
proves that the derived document differs only by final page-derived field text
and internal `layout_page_field` substitution markers. These markers require
source replay and are never emitted by source extraction or accepted as compiler
source input. The original-source digest is also carried in output provenance.
Read-only paragraph policies and original raw XML/package digests are preserved.

Exact `w:pgNumType` decimal section starts (`0` through `999999`) restart PAGE
display text without changing physical pagination or NUMPAGES. Unspecified
starts continue physical page numbering, including parity blanks. A restart
on a page shared by continuous sections refuses rather than guessing ownership.
Chapter numbering attributes, other number formats, malformed starts and display
overflow remain refused. The browser fixture starts at 7 and proves native
`Page 7 of 2` / `Page 8 of 2` in both stories and the page-eight body field.

Simple and qualified flat complex PAGE/NUMPAGES instructions admit only the
optional `\* Arabic` and `\* MERGEFORMAT` switches (each at most once, either
order). Arabic selects decimal digits; MERGEFORMAT preserves the existing
single result run's exact formatting while discarding its cached text, matching
[Microsoft's field-format semantics](https://support.microsoft.com/en-us/word/format-field-results).
Raw instruction bytes remain bound by source digests and field paragraphs stay
read-only. Other formats, CHARFORMAT, numeric pictures, duplicate switches,
locked/dirty fields, nested fields and multi-run results still refuse.

Text-run `w:vertAlign` values `subscript` and `superscript` use an explicit
font-metric simulation profile. Native shaping reads the embedded font's
[OS/2 script size and offset metrics](https://learn.microsoft.com/en-us/typography/opentype/spec/os2),
scales horizontal advances and vertical metrics independently, and carries a
digest-bound `script_transform` on each affected fragment. Page paint applies
the same horizontal/vertical outline scale and baseline offsets; no CSS text
measurement or fixed percentage is used. The output's `font_size_millipoints`
continues to identify the authored run size; the compiled glyph paths are the
rendering authority, not an instruction to re-shape at that size.

This is a deterministic use of the font's recommended simulation metrics, not
a claim that Microsoft Word uses the same policy. Missing/truncated/invalid
OS/2 metrics, non-reducing scales, script paragraph marks, list/note markers,
controls, and combinations with underline or highlighting remain refused.
Direct script paragraphs remain read-only; `baseline` explicitly resets an
inherited vertical alignment. OpenType `sups`/`subs` glyph substitutions, custom
`w:position` offsets, and mathematical equation layout are separate capabilities.

Stored output should first pass `decodeNativeDocxPagePaintV1`, then
`decodeNativeDocxPagePaintForRequestV1` for exact request/page/line/glyph/style
identity joins. When provider-authenticated path geometry and the legitimacy of
a refusal must also be proven, use the async
`validateNativeDocxPagePaintForRequestV1`, which recompiles with the exact
provider and compares the complete wire output. These content hashes detect
projection substitution between engine stages; they are integrity joins, not
external authorization signatures, so the upstream projections must still
remain inside the trusted native pipeline.

See [the native DOCX contract migration guide](../../docs/DOCX-NATIVE-CONTRACT.md)
for scope and rollout boundaries.

Pure document pagination, header/footer tokens, and comment-thread helpers.

```bash
npm install @injoffice/docs
```

```ts
import { DEFAULT_PAGINATION_OPTIONS, paginate } from '@injoffice/docs'

const layout = paginate(
  [
    { id: 'heading', height: 52 },
    { id: 'paragraph', height: 420 },
    { id: 'next-page', height: 80, breakBefore: true },
  ],
  DEFAULT_PAGINATION_OPTIONS,
)
```

The package intentionally contains no editor framework or DOM measurement layer. The legacy `paginate(BlockBox[])` helper remains available unchanged and isolated from native pagination.

## Browser-safe native page output

Use `@injoffice/docs/native-page-paint-output` to validate transported page-paint
output in a browser. This output-only entry exports `decodeNativeDocxPagePaintV1`,
the protocol/version/limits and output types without loading Node font providers,
HarfBuzz, bidi or pagination. It uses the exact same strict output decoder as the
server compiler. Full request/source replay remains a compiler-side operation.
