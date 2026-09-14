# Native DOCX contract and migration

The versioned `injoffice.docx.native` contract is the boundary between the
authoritative Go OPC engine and browser layout/editing code. It establishes a
native model without making `@injoffice/docs`, ProseMirror, DOM nodes, or HTML a
second source of truth.

## Version 1 guarantees

The additive `RunV1.page_field` metadata represents qualified simple decimal
`PAGE` / `NUMPAGES` instructions. Such runs are text-kind with **empty source
text**, not the cached field result, and their paragraph remains read-only.
Consumers with older strict decoders must upgrade together with the Go engine.
The native page-paint compiler emits bounded per-page shaping variants after
body pagination; content-only previews should label unresolved fields rather
than display a guessed number. See the [page-field rendering profile](../packages/docs/README.md)
for the exact supported subset and remaining refusals.

- Every document, story, block, paragraph, run, table row/cell, drawing,
  section, comment, and unsupported record has a durable InjOffice identity.
- Every modeled OOXML object retains a part-qualified source anchor with byte
  extent and XML fingerprint. Anchors locate and guard surgical patches; they
  are not user-facing identities.
- The original package revision and SHA-256 fingerprint support future
  compare-and-swap persistence.
- Unmodeled OPC parts are inventoried with a mandatory
  `preserve-verbatim` policy. Unsupported constructs carry an explicit scope,
  preservation behavior, and refusal message.
- Sections use Word's native twip geometry. Drawings use EMUs and distinguish
  inline from floating placement.
- Paragraphs retain style and numbering references. The table projection is
  intentionally conservative: rows, cells, spans, vertical merges, and cell
  paragraphs are represented; nested or otherwise unsafe table content makes
  the owning object read-only and produces an unsupported record.
- Headers, footers, footnotes, endnotes, and comments are linked through stable,
  typed story/reference identities instead of flattened into body text. Note
  stories retain their owning relationship and content/separator role; native
  note-label runs are typed self-references. Comment
  bodies are full `comment` stories in `comment_stories`; each comment's native
  OOXML id must match the referenced story's `native_story_id`.
- Both TypeScript and Go reject unknown fields, invalid union payloads,
  duplicate identities, wrong-kind or dangling references, malformed anchors,
  and ambiguous read-only policies. Both consume the same golden and shared
  invalid fixtures. Binding-field and structural-constraint tests compare both
  languages to the published JSON Schema; cross-object rules such as typed
  references and anchor containment remain runtime validation rules.
- Canonical encoders sort object keys while preserving array order. Array order
  remains meaningful for stories, blocks, runs, rows, and cells.

## Identity and source-anchor rules

The future Go parser owns identity generation. An identity must remain stable
when unrelated parts or sibling content change. It must not be derived only
from a positional array index. A parser may seed an identity from an existing
OOXML identifier or from a part-qualified structural anchor plus content
fingerprint; once issued, the gateway must carry it through revisions or
publish an explicit identity remap.

`start_byte` and `end_byte` refer to the original uncompressed OPC part bytes,
with an exclusive end offset. A patch must verify `xml_sha256` before using an
anchor. If the anchor or expected package revision is stale, persistence must
refuse rather than search heuristically and mutate a different object.

Every parsed story has a nonempty anchor. Its `part_name` must equal the story's
part; every descendant anchor must use that same part and remain within its
parent byte extent. The body part is exactly `source.main_part`, and section
anchors are contained by the body anchor. OPC part names use a conservative
canonical URI subset: no absolute paths, empty/dot segments, backslashes,
query/fragment syntax, trailing-dot segments, lowercase percent escapes, or
literal/percent-encoded traversal and separators.

## Resource and JSON safety limits

The contract is a gateway boundary, so validation is deliberately bounded:

- raw JSON payload: 8 MiB;
- nesting: 64 levels;
- total traversed values: 100,000;
- any collection: 10,000 items (edit operations: 5);
- one text run: 1,048,576 Unicode characters;
- returned validation issues: 100.

Integers are restricted to JavaScript's safe range. Anchors must have a
strictly positive extent. JSON `null` and negative zero are rejected globally
before typed decoding because either can otherwise collapse into a different
meaning across Go and JavaScript. Browser/gateway callers should decode raw
payloads with `decodeNativeDocxJson`; `decodeNativeDocxDocument` remains useful
for already-parsed in-process values and applies the same traversal, null, and
negative-zero checks.

References are type checked: section starts target top-level body blocks;
header/footer references target the corresponding story collection; note and
comment runs target the correct native story/comment kind; comment metadata
targets a comment story; and drawings target passthrough parts whose content
type is image, audio, or video. Resolving to an unrelated global id is invalid.

## Native Go extractor status

`go/docxpatch.ExtractNativeDocumentV1` now emits this contract directly from
DOCX bytes without Mammoth, HTML, a browser DOM, or whole-document
regeneration. It discovers arbitrary valid main-part locations through root
relationships, accepts Transitional and Strict WordprocessingML, validates OPC
paths/content types/relationships case-insensitively where the standard
requires, and rejects malformed, ambiguous, spoofed, or resource-hostile
packages.

The bounded modeled subset is body and linked story paragraphs/runs, Unicode
text, controls, note/comment references, basic tables and sections, and
conservative inline/floating DrawingML pictures. Exact XML byte ranges hash the
original uncompressed part slices. Non-story parts are inventory-only; complex
markup inside a modeled story generates scoped unsupported/refusal records.

Initial imports may provide an application-owned `DocumentID`. Incremental
imports pass the previous valid contract through `NativeExtractionOptions`:
native story/comment identifiers and unchanged part/fingerprint matches retain
issued IDs when siblings move. Unique Office 2010 `w14:paraId` values from
`00000001` through `7FFFFFFF` seed paragraph identity directly; invalid or
duplicate values remain verbatim source metadata but are excluded from
identity generation. The default initial document seed
depends on the main XML rather than the whole package, so unrelated part
changes do not churn document/object identity.

Extractor limits are 128 MiB compressed package bytes, 256 MiB total expanded
bytes, 64 MiB per part, 16 MiB per XML part, 10,000 ZIP entries, 200:1
per-entry compression (plus 1 MiB slack), 100,000 XML elements, and 64 XML
levels. These precede the JSON contract limits below.

## Resolved layout input status

`go/docxpatch.ResolveNativeDocumentLayoutV1` now derives a separate
`injoffice.docx.resolved-layout` v1 projection from DOCX bytes and the native
extractor. It is intentionally not added to the persisted native document
contract: `NativeDocumentV1` remains the authoritative identity, anchor, and
passthrough boundary, while the resolved projection is deterministic read-only
input for a future shaper and paginator.

The resolver discovers the main part's styles, numbering, theme, and font-table
relationships at arbitrary OPC locations with case-equivalent lookup and
expected content-type verification. Its cascade order is:

1. `docDefaults` paragraph/run properties;
2. default or explicit paragraph `basedOn` chains, base first;
3. ordinary numbering-level paragraph properties;
4. direct paragraph properties from exact anchored story XML;
5. for text runs, paragraph-style run properties, the default or explicit
   character-style chain, then direct run properties.

Style-layer bold, italic, and hidden properties use Word toggle semantics;
direct values are absolute. Missing ancestors and cycles are diagnosed, cyclic
layers are excluded, and any acyclic descendant layer remains explicit rather
than disappearing silently. Partial `numPr` values merge by attribute across
the cascade for direct paragraph formatting. For style-supplied `numPr`, Word's
style `ilvl` is ignored and the concrete level is selected only by one exact,
bounded `w:lvl/w:pStyle` reverse link; missing, duplicate, and multi-level matches
are refused and the selected style ID is digest-bound. Ordinary `abstractNum`/`num` levels, replacement levels,
`startOverride`, level text/suffix/alignment/indent, and marker run formatting
are projected with deterministic source-order counters keyed by concrete `numId`, restart state,
literal percent text around `%1` through `%9` placeholders, and exact text-margin-anchor geometry.
Valid `singleLevel`, `multilevel`, and `hybridMultilevel` values are authoring metadata; actual
declared levels remain authoritative. Replacement-level `start`/`lvlRestart` values do not alter
counters, while concrete `startOverride`, abstract `lvlRestart` values from 0 through 7, and
numbering `w:tab w:val="num"` stops are applied and attested. Unknown or foreign
numbering-root children remain preserved and silent while numbering is unused;
because their semantics may be global, the resolver emits a blocking diagnostic
for every concrete `numId` reference instead of erasing or guessing them. The numbering source
attests the owning relationship part, relationship identity/target, raw
numbering part/content type, both raw SHA-256 hashes, and a canonical digest of
the resolved marker sequence. Definition/model digests use the versioned
`injoffice.canonical-json/utf8-v1` serializer in both Go and TypeScript, including
identical HTML-character, Unicode-separator, astral, and non-ASCII handling.

The supported resolved subset includes explicit Latin font names, half-point
sizes, RGB colors, language, run RTL/hidden state, valid Word highlights and a
conservative underline subset; paragraph bidi, physical and logical indents,
alignment, spacing/line rules, keep flags, page breaks, and widow control; and
ordinary decimal, letter, Roman, and bullet markers. Source part names, durable
native object IDs, and scoped preservation diagnostics remain in the derived
model. Encoding validates bounded collections/strings, part names, identity
references, property domains, and duplicate-free style traces.

Exact `w:color/@w:themeColor` values resolve through the theme part's
`a:clrScheme` when that slot is a leaf `a:srgbClr` or an `a:sysClr` with an
exact six-digit `lastClr` snapshot and no DrawingML transforms. Cached `w:val`
is ignored while themeColor is present; `themeTint`/`themeShade`, `auto`,
`a:schemeClr`, missing/duplicate slots, and unknown theme names stay
`THEME_COLOR_PRESERVED`. The extractor projects the same sRGB onto modeled run
colors, table-border `color_rgb`, and clear cell fills (`w:shd/@w:themeFill`)
and accepts identity `w:space="0"` on single/none borders. Exact
`asciiTheme`/`hAnsiTheme` latin slots resolve to the theme `a:latin/@typeface`
name; page paint still requires that face's attested package bytes and never
invents Calibri. East-Asia or complex-script selection, font hints, automatic/character-unit
spacing, picture bullets; legal, ordinal, cardinal/text number formats;
ambiguous or malformed placeholders/overrides; tracked numbering changes;
conditional table-style region effects, or unknown extension semantics remain
unresolved. Simple table styles without `tblStylePr` may project whole-table
borders and clear fills onto the existing page-paint commands. Those constructs stay in the original package and produce explicit
`preserve-verbatim` diagnostics only when their concrete `numId` and level are
referenced; structural parser and collection limits remain hard. Every accepted
twip is bounded to `180143985094819` before its exact ×50 milli-point
projection; raw negative zero remains invalid. The resolver does not use
Mammoth, HTML, browser/DOM APIs, or a lossy parse-regenerate path.

## Native shaping and line-breaking status

`@injoffice/docs` now consumes the persisted native contract and separate
resolved-layout projection through a versioned
`injoffice.docx.shaping-request` v1 boundary. The TypeScript resolved-layout
binding mirrors the exported Go JSON fields and rejects unknown fields,
unbounded traversal/collections, null and negative-zero hazards, unsafe OPC
part names, invalid property domains, and broken paragraph/run/table identity
joins before any font provider runs.

`shapeNativeDocxLinesV1` accepts explicit integer milli-point width and tab
constraints and injects the shared `@injoffice/font-metrics/layout`
resolver/shaper interfaces. It converts twips and half-points without device
pixels, segments authored text into explicit OpenType
script/language/direction requests, validates returned font resources and
cluster/glyph ranges, verifies loaded bytes against the manifest-backed face
SHA-256, and emits deterministic
`injoffice.docx.shaped-lines` v1 lines/fragments tied to native paragraph/run
IDs and UTF-16 source offsets. Output also binds the exact font-manifest ID and
revision plus resolver/shaper IDs and revisions, making cache provenance
explicit. The validated wire request, manifest, provider identities, and bound
callable references are snapshotted before the first provider call; live caller
mutation cannot alter the authoritative result, and mid-call provider ID or
revision changes fail closed.

The bounded core preserves authored whitespace, paragraph-content-relative
tabs, OOXML line breaks, supplementary Unicode clusters, non-breaking spaces,
and safe cluster boundaries. It applies resolved physical/logical indents,
first-line/hanging offsets, alignment, before/after spacing, and Word
auto/exact/at-least line heights. Ordinary decimal/letter/Roman/bullet markers
consume the resolver's source-ordered counter vector, final text, and hanging
geometry; TypeScript never reconstructs list state. A skipped table in a
numbered document atomically empties the shaped paragraph projection rather
than silently presenting a partial list.
Resolved paragraph-mark/default run properties provide real font metrics for
blank, hidden-only, control-only, and explicit empty lines without synthetic
output fragments.

Resolved-layout diagnostics fail shaping closed by default. Only an enumerated
paint-only subset may retain advances, and the original diagnostic code/message
is propagated. Provider metadata, coordinates, line heights, block advances,
and output collections are explicitly bounded. Any blocking run, marker, or
provider result atomically refuses the paragraph rather than emitting surviving
run fragments as a misleading partial line. Resource bounds are 64 MiB per
font, 128 MiB cumulative unique bytes, 256 cached/loaded faces, 50,000
resolve/shape calls, and 256 loads per shaping request. Resources are copied,
digest-checked once, and cached by the full exact resolved-face identity.
Runs, faces, decisions, attempted-face IDs, and provider results cross async
boundaries only as owned snapshots. The shaper receives a separate per-face
resource copy (also cumulatively bounded to 128 MiB), never the authoritative
cache. JavaScript cannot freeze individual `Uint8Array` elements without a
per-shape full-font copy, so consuming the isolated provider-facing bytes is an
explicit trusted-shaper responsibility; manifest/face/metrics/run/provenance
metadata remains immutable and cannot be rewritten by the provider.
The Node 22 `@injoffice/font-metrics/harfbuzz` entry now provides a genuine
HarfBuzz qualification boundary for the qualified page-paint slice. It verifies
the exact `harfbuzzjs@1.6.0` JavaScript entry/loader/manifest and shaping WASM
plus the HarfBuzz 14.3.0 runtime, applies a bounded structural preflight to
digest-bound sfnt/TTC bytes and the v1-qualified metric projection, and emits
real glyph IDs, positions, and monotone UTF-16 clusters. Its `providerRevision`
is a SHA-256 over runtime artifacts, generated Unicode 13 data, configuration, and
source provenance. The shaper remains resolver-neutral and does not duplicate
the separate inventory/resolution or outline-provider lanes.

The provider accepts only complete non-system fixed TrueType horizontal
Latin/Common, Arabic, and Hebrew runs with zero synthetic spacing and explicitly
qualified `kern`/`liga` settings. The separate source-digest-verified
`bidi-js@1.0.3` boundary applies Unicode 13.0.0 UAX #9 paragraph resolution,
explicit paragraph/run direction, cluster-aligned L1/L2 visual ordering, and
logical-to-visual maps. Because bidi-js 1.0.3 classifies UTF-16 code units, the
boundary atomically refuses every supplementary scalar before resolution. A
generated classifier from vendored, digest-verified Unicode 13 sources supplies
the shared Script/Common/Inherited, whitespace, punctuation,
Default_Ignorable/control, and assigned-repertoire decisions without host
Unicode or locale authority.
Soft-wrapped `both` lines deterministically expand paintable U+0020 clusters in
visual order; final and hard-break lines use logical start alignment without
expansion. Variations, CFF/CFF2, authored bidi controls, vertical directions,
unadvertised explicit feature settings, structurally malformed fonts, missing
glyphs, default ignorables, unsupported scripts, distributed character spacing,
and justified lines without a qualified U+0020 opportunity fail closed.
Table grids, drawings, pagination controls,
note/comment display text, unresolved theme tint/auto colors, missing sysClr lastClr snapshots, and
script-dependent gaps remain explicit scoped diagnostics. This is input to the native paginator/painter, not page
placement or rendering itself, and imports no Mammoth, HTML, DOM, canvas,
React, or Konva.

For the bounded numbering slice, the shaper consumes only Go-resolved marker
text/counters. It shapes marker glyphs through the same attested HarfBuzz lane,
positions left/right/start/end/center labels from shaped advance, and resolves
suffix tabs against the hanging stop or explicit default interval. Marker geometry
is pre-refused when shaped advance crosses the label end, and an explicit numbering
tab at or before the shaped marker end refuses the paragraph atomically. A numbered
paragraph refusal, skipped numbered table, source-hash mismatch, or marker
resource failure empties the complete shaped paragraph projection. Pagination
independently recomputes the marker end from the shaped visible marker fragments,
then recomputes `text_start_millipoints` and any tab-suffix target from the resolved
suffix, exact numbering tab, and attested Word default-tab interval. Thus even a
self-consistent tamper of the shaped marker projection is refused. Pagination and
page paint exact-join marker IDs, definitions, counters, text, geometry,
font properties, source hashes, and the model digest; page paint replays only
already-shaped marker glyphs in canonical visual order. Mere numbering-part provenance does
not turn an unrelated unnumbered drawing or table refusal into a numbering-wide
erase; that gate requires a resolved marker or a concrete numbering diagnostic in
the affected scope.

The wrapper is intentionally an invocation-scoped, bounded per-compilation
provider rather than a reusable long-lived service; every native compilation
gets hard face/byte/call budgets and an atomic result.

## Native ProseMirror transaction adapter v1 status

`@injoffice/docs` now exposes a canonical, fail-closed
`injoffice.docx.prosemirror-transaction` v1 adapter. It consumes the validated
native document plus a complete before/after body projection and a maximum of
five exact successive replace steps. Projection positions are UTF-16 code
units; durable paragraph/run identities and run XML fingerprints remain the
native join. The adapter emits only existing `run` text replacements inside
the shared `injoffice.office.mutations` v1 envelope, with the exact
`source.package_sha256` as save authority. Multiple ordered changes to one run
are coalesced; Go remains the only OOXML mutator and preservation authority.
The shared typed mutation envelope additionally admits Go's proven
`paragraph` selector only for a paragraph containing exactly one run, which
must be text. It uses the paragraph ID and paragraph XML fingerprint, and cannot overlap the
underlying run selector in one batch. The ProseMirror adapter itself continues
to emit run selectors because its projection is run-addressed.

The adapter rejects tables, non-text inline nodes, drawings, notes/comments,
fields, tracked or unsupported native markup, stale identities/revisions,
schema or mapping drift, open/structural slices, changed marks/properties,
cross-run edits, overlap/reordering, surrogate splits, unsafe XML characters,
unattested edge whitespace, partial projection coverage, hostile objects/unknown
keys, no-ops, and resource overflow. Every refusal contains issues and no
payload or envelope. Canonical output stays within the shared 3 MiB payload
and 4 MiB envelope limits and never invokes HTML, DOM parsing, browser layout,
Mammoth, OOXML reconstruction, or package regeneration.

The optional host projector is intentionally only an integration seam. The
website TipTap schema still needs to carry durable native IDs/anchors and emit
the exact projection at save time. ProseMirror collaboration sequence/rebase
state must remain separate from exact package-byte CAS and cannot become DOCX
rendering or persistence authority.

## Native pagination v1 status

`@injoffice/docs` now exposes `injoffice.docx.pagination-request` v1 and
`paginateNativeDocxV1`. The request binds one validated native document,
resolved-layout input, shaped-lines output, and extractor-owned
`injoffice.docx.pagination-settings` v1 attestation to the same document ID and
revision. The attestation relationship-discovers the settings part and binds
the package fingerprint, actual owning `.rels` part/fingerprint/relationship
ID, and settings part/fingerprint. It distinguishes absent Word defaults from a
bounded modern profile and carries explicit refusal diagnostics for unknown
layout-affecting semantics. A dedicated shaped-lines decoder verifies exact keys, bounded
traversal, deterministic paragraph/line/fragment IDs, block-height arithmetic,
provider provenance fields, and native story/run joins before page placement.
Accepted settings nodes have exact attribute, child, and whitespace-only text
shapes with duplicate-singleton detection in both Transitional and Strict
packages. Theme-font language, shape defaults, attached templates, native math
defaults, placeholder/revision display settings, enabled field-result updates,
and other settings not proven irrelevant to the resolved/shaped input remain
explicit refusals rather than a broad neutral allowlist. `updateFields=false`
is attested explicitly, and content-type identity uses ASCII-only case folding.

The supported projection is intentionally narrow: source-ordered body
paragraphs in exact single-column or qualified equal-width column grids whose
column width equals the explicit shaping width. Twips convert exactly to
integer milli-points at 50:1. Page
width/height, physical margins, gutter, and body box remain explicit. The
paginator handles continuous, next-page, even-page, odd-page, and next-column
section transitions. Continuous transitions share the current page only for
an identical single-column physical grid and identical header/footer references.
Next-column transitions require the same exact physical grid, advance to the
next column, and start a new physical page when the final column is exhausted.
parity insertion produces a real header/footer-free blank page owned by the
preceding section, with an explicit reference to the section it precedes. The
first section's transition type is ignored because there is no preceding section.

WordprocessingML equal-width columns require an explicit count and spacing for
multi-column pagination, an exactly divisible body width, and at most 45
columns. Explicit `w:col` geometry is accepted only when every width and each
authored gap is valid (omitted `w:col/@space` has the standard zero value), the
final gap is zero, total geometry exactly covers the body, and all
column widths are equal. Section and column IDs are carried through paginated
pages, paragraph slices, placed lines, canonical hashes, and page-paint output.
Terminal balancing accepts uniform-height, zero-spacing lines, including multiline
paragraphs, without cross-paragraph break constraints. A deterministic
quotient/remainder plan gives earlier columns at most one extra line. Multiline
splits must respect `keep_lines` and default-on `widow_control`; when the ideal
plan violates either, pagination refuses atomically instead of choosing a
constrained rebalance. Source line ordinals and paragraph continuation flags
remain explicit across columns and pages. Unequal widths and general
constrained balancing remain outside this column slice.

This bounded interpretation follows the OOXML definitions of
[`noColumnBalance`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.nocolumnbalance?view=openxml-3.0.1)
and [`widowControl`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.widowcontrol?view=openxml-3.0.1).
It does not establish visual equivalence with Microsoft Word.

Paragraph placement applies `page_break_before`, `keep_lines`, and a provable
atomic subset of `keep_next`. A keep-with-next chain is placed only when every
multiline member is also `keep_lines`; otherwise pagination refuses because
boundary-aware splitting is not implemented and `keep_next` alone does not
make the whole paragraph indivisible. Adjacent after/before spacing uses Word's maximum-spacing rule.
Paragraph-before spacing at an empty page is honored only on the first content
page of a section; it is suppressed after automatic, explicit, keep, or widow
page movement. Keep chains are planned in a single reverse pass, so the bounded
paragraph/line input cannot induce quadratic chain rescans.
Widow/orphan control defaults on when absent and prevents singleton first/last
page lines where a valid split exists. An unsatisfiable keep or widow policy
refuses the complete projection instead of silently overriding Word semantics.
Every content page contains deterministic paragraph slices and placed lines;
their IDs derive from durable native paragraph and shaped-line identities, not
physical page numbers. Physical and per-section ordinals remain explicit, and
all coordinates stay in integer Office units. Wire-visible pagination and
page-paint diagnostics use explicit UTF-16 code-unit ordering rather than
host-locale collation.

Unequal-column paragraph flow is a separate bounded profile: exactly two
explicit unequal widths in one section, with source-attested
`compat/noColumnBalance=true`. The exact OnOff leaf is optional in settings;
missing `val` means true, explicit false retains existing behavior, and true
outside this qualified profile refuses. Duplicate, nested, spoofed, or foreign
attributes do not create an attestation. The ordinary geometry helper still
refuses unequal widths unless its explicit opt-in is used by the qualified flow.

This profile supports nonempty ASCII LTR text paragraphs, zero spacing and
indents, natural line height, no notes/tables/fields/drawings/numbering or
headers/footers. Multiline paragraphs require `keep_lines=true` at both widths;
all paragraphs must fit an empty column at both widths. `keep_next` and
`page_break_before` remain excluded. Input is bounded to 256 paragraphs, 256
lines per paragraph per width, and 100,000 source UTF-16 units, in addition to
existing wire, shaping, page, and placement budgets.

The optional pagination-request `column_shaped_lines` pair carries both complete
width-specific snapshots. Request validation joins each to source and resolved
layout, proves complete monotone text coverage, shared font/provider provenance,
per-run font metrics, per-source-position script identity, and natural line
geometry. Deterministic replay first tests a whole paragraph at the current
column width; if it does not fit the remaining height, replay selects the next
column's complete paragraph. It never splits a paragraph across widths. Last
columns may remain unused. The selected `shaped_lines` snapshot must equal this
plan, and page-paint hashing includes both candidates. No source diagnostic is
removed: only the exact section-local `UNEQUAL_SECTION_COLUMNS` record is
discharged after complete qualification. Unsupported or exhausted placements
produce no partial pages.

This implements the sequential fill policy described by
[`noColumnBalance`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.nocolumnbalance?view=openxml-3.0.1)
and the paragraph atomicity of
[`keepLines`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.keeplines?view=openxml-3.0.1).
The imported Go fixture and actual HarfBuzz outline/paint replay validate this
pipeline; they are synthetic OOXML evidence, not a Word visual reference.

The separate read-only rectangle preview now retains a bounded floating
textbox's page-relative source offsets in `textbox_geometry.items[].page_anchor`.
It requires exact `wp:anchor` offsets from the page, `wrapNone`, zero distances,
zero effects and relative height, foreground/unlocked placement, layout-in-cell
and overlap enabled. Nonnegative offsets are bounded to 127,000,000 EMU and
must convert exactly to millipoints. Container and ordered offset-leaf anchors
bind this metadata into the existing local shape-paint digest. Inline previews
remain unchanged. This is source placement evidence only: physical-page
assignment, page containment and composition with body text remain unfinished,
and native drawing diagnostics remain intact. The complete source profile is
documented in [the Docs package](../packages/docs/README.md#explicit-rectangle-textbox-preview).

Whole-footnote body reflow admits one or two footnotes, each containing 1–16 paragraphs, in a
single-section, single-column document. Multiple note paragraphs require a complete
resolved `keep_next=true` chain on every nonfinal paragraph, with final keep-next
absent or false and the sole owning label in the first paragraph. Every multiline
member requires `keep_lines=true`; oversized reference/group pairs remain refused.
This follows authored [KeepNext](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.keepnext)
and [KeepLines](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.keeplines)
semantics without implementing their oversized-group fallback. Body paragraphs and the note use simple
LTR text, natural line heights and zero spacing/indents. Every multiline body
or note paragraph must have resolved `keep_lines=true`; body keep-with-next, explicit
page breaks, tables, drawings, fields, numbering, headers/footers and explicit
note positions are excluded. The exact instruction-only separator also has zero
spacing and one natural line. Each footnote requires a unique reference in a
distinct body paragraph. Its reference paragraph plus individual separator/note
area must fit an empty page; otherwise the whole projection refuses. Decimal
numbering follows body reference order, including when note2 moves to a new page.

Reservation and final placement share the same note geometry implementation.
Before placing a reference paragraph, the paginator measures the tentative page
note group in source order with one shared ordinary separator. It compares used
body height plus the whole reference paragraph and the total group area with
physical page height; the existing reservation is replaced, not counted twice.
If the total does not fit, only the new reference and its note advance, leaving
previous body paragraphs and notes unchanged. A combined two-note group may
exceed one page when each individual reference/note pair fits separately.
Subsequent body paragraphs use the reduced space only on that page; advancing
restores full body height. Final placement must match each page's ordered source
note/reference IDs, global decimal numbers, single separator and measured area,
with no unreserved groups. Source, label, relationship and exact diagnostic
validation remain mandatory. Deterministic request replay executes the same
flow; no caller-supplied reservation or source mutation is accepted.
Same-paragraph multiple references, more than two-note reservation, general
footnote continuation and oversized individual pairs remain outside this slice.

The existing omitted document-level footnote position remains page bottom, as
described by [`FootnotePosition`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.footnoteposition?view=openxml-3.0.1).
Explicit document/section positioning remains refused. Synthetic imported
fixtures verify reference-stays/later-body-moves, reference-and-note-move,
oversized-pair refusal and explicit-position refusal through actual native font
shaping and outline paint. They do not establish Word visual equivalence.

The qualified note slice is relationship-resolved and source-order driven.
Each referenced footnote/endnote content story contains paragraphs only,
exactly one self-label, and exactly one matching body reference. Labels use
independent decimal sequences per kind, beginning at one. A used kind has one
exact separator story. The paginator atomically reserves the unused bottom of
each referencing page for that page's footnotes. It atomically places all
endnotes at the end of the document. Whole fitting groups may move to one fresh
final page; the bounded continuation profile below can append additional pages.
Placed note stories preserve kind, native ID, relationship ID, role, reference
run, assigned label, geometry, section ID, column ID/ordinal, and shaped-line
provenance. WordprocessingML note-reference superscript remains in OOXML, but
native shaping and page paint refuse vertical alignment until its scale,
baseline, and advance metrics have a qualified authority.

Semantic refusal is distinct from malformed input. Invalid wire input returns
validation issues. A valid request containing unsupported pagination semantics
returns a versioned `status: refused` result with diagnostics and empty page and
section arrays, so no partial layout can be mistaken for complete output. The
strict `decodeNativeDocxPaginatedLayout` decoder verifies output union
exclusivity, deterministic IDs, section/page and slice/line joins, contiguous
line geometry, exact one-slice coverage of every placed line, globally
contiguous paragraph slice ordinals/ranges/continuation flags, physical source
ordering without overlaps, page bounds, provenance, null/negative-zero safety,
and global resource limits. Parity fillers must immediately precede a section
whose even/odd break type and resulting one-based page parity match the filler.
`decodeNativeDocxPaginatedLayoutForRequest` adds the source join a standalone
output cannot infer: every shaped body paragraph/line must be placed exactly
once in native source order, and source sections, starts, breaks, geometry,
header/footer references, complete provenance, vertical placement, paragraph
spacing, and deterministic page/slice break choices must equal the exact
request. `paginateNativeDocxV1` runs this same check before returning paginated
output.

V1 refuses tables outside the exact qualified fixed-layout slice,
floating/anchored/wrapped drawings and every inline drawing outside the exact
page-paint image slice, body comments and non-note references, note-bearing
pages with multiple columns or section grids, column separators, unequal or
incomplete explicit columns, ambiguous/defaulted multi-column spacing,
continuous multi-column transitions, incompatible shared-page geometry,
varying section body widths that require re-shaping, and legacy/unknown settings.xml
compatibility, mirror-margin and top-gutter semantics, inline page/column controls,
drawing/comment/non-note reference advances absent from shaped output, and all unknown
pagination-affecting source records. Pagination retains explicit header/footer
references as provenance; the downstream native header/footer planner uses the
extractor-owned `title_page` policy, the attested even/odd setting, physical
page parity, and per-kind section inheritance to select first/even/default
stories. An absent settings part attests Word's 720-twip tab default but
also the omitted compatibility setting's mode-12 behavior, which pagination v1
refuses. Supported modern settings bind an explicit tab stop and compatibility mode 15,
and the paginator verifies that the shaped tab interval matches. The extractor's explicit
schema-default section is accepted only when its complete standard geometry is
present; other invalid or inferred geometry refuses.

The exact inline-image slice admits only an internally related, preserved,
digest-bound static PNG with exact preserved byte length, identity DrawingML
transform, no crop/effect/wrap,
and positive EMU extents exactly divisible through the reduced `10/127`
milli-point conversion. The line formatter treats that picture as a glyphless
atom whose width and ascent are the qualified extent. Page paint emits one
content-addressed media resource and one explicit full-crop, identity-transform
command at the shaped baseline. Missing/duplicate/extra assets, relationship or
media digest drift, malformed/animated/oversized PNGs, and unsupported geometry
refuse atomically with no pages or resources.

The v1 media inventory is intentionally body-story scoped. A picture in a
selected header or footer is diagnosed and refuses the complete paint result;
it is never omitted while body pages are published. Story-scoped asset
selection and relationship closure require a later contract version.

Notes requiring body reflow or continuation outside the profiles below
refuse atomically. So do custom
numbering/restarts/positions, ambiguous/duplicate/missing references, unpaired
labels, duplicate IDs or relationship drift, cycles, nested tables, drawings,
fields, and unknown note markup. Exact Word `w:separator` and
`w:continuationSeparator` instruction leaves are admitted only in their
matching `-1` and `0` sentinel stories. Page paint derives one bounded black
separator rule from the placed ordinary sentinel. A continuation rule spans the
full single-column text width. A dormant continuation sentinel is not shaped;
the compiler performs one bounded shaping retry only after pagination emits
`note-continuation-shaping-required` for its exact source story.

Endnote continuation admits exactly one content endnote in one single-column
section with no footnotes. Every note paragraph must have zero before/after
spacing, no `keep_next`, no `page_break_before`, and no fields. Multiline
paragraphs with resolved `keep_lines=true` remain indivisible. Other paragraphs
may split at exact shaped-line boundaries. Default/enabled widow control
requires at least two lines on each side of an internal split; explicit false
permits one-line slices. The paginator greedily fills the available area and
backs up to a legal boundary, adding a page when the next line group plus its
separator cannot fit. An unsatisfiable keep/widow group refuses atomically. The first slice uses the ordinary
separator even when moved to a fresh page; only later slices of the same note
use the source-bound continuation sentinel. Custom continuation notices remain
unsupported source semantics.

Footnote continuation admits multiple content footnotes in one single-column,
single-section paragraph body with modern settings and no header/footer stories.
The joint flow profile requires zero-spacing left-to-right body/note paragraphs,
no numbering, fields, drawings, run controls or `keep_next`. Body
`page_break_before` remains effective. Existing whole-note reservations retain
their separate authored keep-chain profile.

The planner consumes body lines and footnotes together. It reserves a legal
note prefix before admitting subsequent body lines and honors body and note
keep/widow groups. Every new note retains its first legal slice on its reference
page. If an earlier carried note prevents the next reference group fitting, a
note-only page advances the carried note before retrying the body group. Notes
remain in reference order: an earlier note finishes before a later note begins.
A carried note may finish beside a newly referenced note under one continuation
separator. A page without carried content uses the ordinary separator. Once the
notes finish, later body pages regain their full height.

Every footnote slice is aligned to the page body bottom. The source label appears
only in its original paragraph. Activated separator shaping remains lazy. The
final note placement recomputes the plan from the actual source and checks its
body page/line assignments and measured reserved height before committing.
Missing or unsupported activated separators, unsatisfiable line groups, budget
exhaustion, and failed replay discard all pages. The note limit is the existing
10,000 notes, with at most 100,000 total placed body/note/separator lines and
2,048 pages; the existing request traversal budget also applies. Mixed footnotes/endnotes, note tables and unsupported body profiles
remain outside this joint flow profile.

Slices retain the existing page/story placement IDs and full source line
identities. Content lines and the original label occur exactly once across
all pages; no label, paragraph, or separator is synthesized. Request-bound
validation replays the complete placement to reject omitted, reordered,
duplicated, or transplanted slices. Geometry, font provenance, page budgets,
and activated sentinel diagnostics remain mandatory. Placement is staged and
any failure discards the entire result.

This profile follows the OOXML definitions of
[`keepLines`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.keeplines?view=openxml-3.0.1),
[`continuationSeparator`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.continuationseparatormark?view=openxml-3.0.1),
and document-end [`endnote pos`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.endnoteposition?view=openxml-3.0.1).
It does not establish visual equivalence with Microsoft Word.

Pagination limits are 2,048 pages, 100,000 placed lines, 100,000 paragraph
slices, 1,000 diagnostics, two million traversed output values, and one
trillion milli-points for any coordinate or checked sum. No DOM, HTML, CSS
layout, canvas measurement, browser renderer, Mammoth, or legacy pixel
paginator participates. Header/footer selection and static paragraph placement
are performed only after complete body pagination; table and footnote layout
are integrated with qualified table and note placement before complete page
paint is exposed.

## Native header/footer page paint

`injoffice.docx.header-footer-layout` v1 resolves each header/footer reference
kind independently across sections. A missing kind inherits the preceding
section's effective reference; a missing kind in the first section is blank.
The first content page selects `first` only when exact `w:titlePg` is enabled,
otherwise an even physical page selects `even` only when the attested settings
enable distinct even/odd stories, and all remaining pages select `default`.
Parity filler pages remain blank.

Selected static paragraph stories reuse the native resolved-layout and shaped
lines projections. Header origins derive from `header_twips`; footers are
bottom-positioned from `footer_twips`. The complete story must fit between that
edge distance and the body box, with horizontal placement constrained to the
exact section body width. Relationship identities, selected story kinds,
source diagnostics, paragraph properties, line geometry, band bounds, and
resource limits are checked before any page is exposed. Exact embedded inline
PNG pictures in selected header/footer stories reuse the same
`paint_inline_image` command, package-part digest join, and EMU geometry as
body pictures. Tables, shapes, non-PNG drawings, note/comment references, bidi,
justification, pagination controls, and ambiguous relationships refuse the
complete plan.

The planner's canonical SHA-256 sorts object keys by code unit while preserving
semantic array order. Page-paint provenance binds that hash, and request-bound
validation deterministically recomputes the full plan. Static selected stories
then use the same content-addressed glyph-outline pipeline as body lines.

Native field instructions are still preserve-only. In particular, a cached
visible result is never treated as an authoritative `PAGE` value: selected
stories containing `PAGE`, `NUMPAGES`, section-numbering fields, switches, or
malformed/nested fields refuse atomically. Exact PAGE support requires a future
page-specific field/shaping contract that also attests the absence of
`w:pgNumType` restart/format ambiguity.

All request provenance checks are exact structural joins, not signatures or
authorization. A host must keep the extractor-owned document, resolver output,
shaper output, and settings attestation within one trusted native pipeline.
Transporting projections independently still requires the complete request
plus request-bound recomputation. Page paint additionally binds canonical
font-manifest, shaped-lines, and complete paginated-layout SHA-256 values.

The canonical font prerequisite is a separate
`injoffice.docx.font-inventory` v1 projection. It preserves font-table and font
parts as untouched package assets while binding exact package, relationship,
part, stored-byte, decoded-byte, face-slot, obfuscation, subset, and OS/2
embedding-rights evidence. Its native-text projection contains document-backed
content digests only and no fallback chains. Resolution reconstructs that
inventory from the authoritative package bytes; system lookup, substitution,
unsigned paths, malformed fonts, and stale manifests are refusal outcomes.
This stage does not parse or serialize glyph outlines.

The Node 22 page-paint prepare boundary carries the exact canonical Go JSON
for that inventory; it does not accept a separate caller font manifest or
resolver revision. Preparation exact-joins document/package/revision/main-part
identity, resolved font-reference coverage, font-table and relationship hashes
and targets, face slots, unique resources, manifest and inventory digests,
collection indices, and licensing flags. Its asset wire carries identities
copied by `ResolveNativeDOCXPagePaintFontAssetsV1`; Node rechecks decoded and
re-obfuscated byte digests plus OS/2 embedding permissions before HarfBuzz is
called. It then resolve/loads every unique authored reference through that
attested embedded-font provider and accepts only a HarfBuzz shaper instance
constructed by the exact pinned runtime module before bidi resolution can run.
Unknown, duplicate, incomplete, stale, substituted, or partially
covered inputs atomically refuse without page-paint output.

## Remaining deliberate non-goals

The native foundation, style-resolution, and shaping-line lanes do not
implement:

- exhaustive DOCX parsing or a model-to-OOXML regenerator;
- unqualified tables, unequal-width columns outside the bounded paragraph profile, column separators, general
  continuous multi-column balancing, dynamic header/footer fields, general
  note pagination outside the bounded native subset, distributed-character
  justification, or broader page painting;
- a ProseMirror schema or transaction adapter;
- structural block/table/drawing patch operations;
- replacement of the current Mammoth view-only fallback.

Adding one of those capabilities requires its own worktree, PR, CI pass, and
merge. A new capability must extend the model through an additive compatible
change or a new protocol version; it must never hide unsupported XML in an
untyped property bag.

## Migration sequence

1. **Implemented foundation:** `go/docxpatch` emits v1 identities, exact
   anchors, linked stories, passthrough inventories, and explicit unsupported
   records. Continue expanding the Office compatibility corpus without
   weakening refusal behavior.
2. Add revision-safe gateway parse/save endpoints. The gateway validates this
   contract before returning or accepting it and keeps original DOCX bytes as
   the persistence base.
3. **Implemented layout-input, exact native numbering, line, bounded body-pagination, and narrow page-paint foundation:**
   resolve the conservative native style/numbering subset in Go, consume it
   through the shared font resolver/shaper, then place the supported
   exact body subset through the renderer-neutral paginator. The
   bounded Unicode-13 visual-bidi and U+0020 `both`-justification slice is
   implemented for the native painter. Next add footnotes, broader column
   balancing and justification, and broader painting stages on these
   versioned outputs; do not
   reinterpret the package through HTML.
4. Translate ProseMirror transactions into versioned native mutation
   operations addressed by durable IDs. ProseMirror remains an input surface,
   not the layout or OOXML model.
5. Enable native DOCX rendering and editing behind a rollout flag; compare
   visual, preservation, editability, and performance results against the
   Office compatibility corpus.
6. Remove Mammoth/HTML interpretation only after the native path passes the
   preservation, Office reopen, fidelity, and performance gates in
   [the independent-engine ADR](ADR-INDEPENDENT-ENGINES.md).

The old and new renderers must not receive parallel fidelity features during
migration. New DOCX compatibility work belongs in the native model, layout,
or surgical-persistence lanes.

### Composed textbox page preview

The Node `renderNativeDocxTextboxPagePreviewV1` API accepts the ordinary complete
page-paint preparation input plus same-source rectangle geometry and supplied
font bytes. It projects only one qualified `PICTURE_GRAPHIC_REQUIRED` diagnostic
out of an internal body copy. The original drawing must precede every modeled
run in its direct body paragraph and occupy its first run's first drawing.
The ordinary body paginator assigns the paragraph's first line to a physical
page; source page offsets then place the rectangle on that page, including its
stroke in the containment check. Other drawing diagnostics, incomplete geometry,
body refusal and shape/font/placement refusal prevent the entire composition.

The distinct read-only approximate output retains original diagnostics and a
canonical source digest. Browser decoding joins body provenance, geometry,
font digest and placement back to the current inspected source. The local
rectangle evidence alone still grants no page layout or mutation capability.

The plural `renderNativeDocxTextboxPagesPreviewV2` API uses the same source and
page-offset policy for 1–64 rectangles. It requires complete geometry and one
exact font per evidence item. Multiple drawings may precede all modeled text
in an owner paragraph. Every selected drawing restriction is removed only from
the internal body copy; any unselected drawing prevents composition. Version 2
contains a source-ordered `textboxes` array, with each owner’s first body line
selecting its page. Browser decoding requires every source rectangle exactly
once, its font digest, coordinates and page containment. Caller inputs are
snapshotted before asynchronous body outlining. The helper uses version 2 for
multiple rectangles, while retaining version 1 for a single rectangle.

Relative position evidence uses policy `relative-position-no-wrap-v2`, retains
signed `x_emu`/`y_emu`, and adds `horizontal_relative` / `vertical_relative`.
Optional `horizontal_align` / `vertical_align` replace the corresponding offset
(which must be zero) and bind the exact `wp:align` source leaf. Accepted axes are
page/margin/column/character horizontally and page/margin/paragraph/line
vertically. Horizontal alignment is left/center/right for page/margin/column;
vertical alignment is top/center/bottom for page/margin. Other combinations
refuse. Physical page containment uses the rectangle edges and clips outward
stroke at page edges; old page-offset evidence retains its full-stroke check.
Source margins come from the owning section, not a note-reduced body box.
The first placed owner line resolves character, paragraph and line origins;
column placement uses that line's section-qualified column. Original drawing
restrictions and source bytes are preserved. This does not add body wrapping,
inline placement, later-run anchoring, or parity-relative margins.

Nondefault textbox stacking is carried in the source-container-bound `stacking`
record (`behind_doc`: boolean, `relative_height`: unsigned 32-bit integer).
Its presence routes page-offset shapes through the version 2 position policy;
the original version 1 API continues to refuse them. Composition preserves the
source-ordered array and validates exact stacking records. The browser partitions
behind/foreground shapes around body paint and sorts each partition by relative
height with stable source-order ties. Changing, omitting or adding a stacking
record invalidates the entire composition. See [the remaining delivery
plan](DOCX-TEXTBOX-LAYOUT-PLAN.md).
