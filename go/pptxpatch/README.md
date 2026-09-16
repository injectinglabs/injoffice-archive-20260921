# pptxpatch

`pptxpatch` reads and writes a bounded, editable PowerPoint OOXML model.

`NativePPTXExtractOptions.AllowInheritedTextPreview` separately enables a read-only
`source-latin-inheritance-approximate-v1` preview. This declared policy combines
presentation defaults, non-placeholder shape master `otherStyle`, shape `fontRef`,
local list styles, and direct paragraph/run properties, in that order. It is not
a qualification of PowerPoint's general inheritance precedence. Each layer is
validated before projection; font, size and color must come from source. Missing
bold/italic resolve to false, and source paragraph defaults supply left alignment
and no bullet only when not authored. The bounded profile accepts graphic ASCII
Latin text, disables authored kerning, and omits validated terminal language and
checking metadata from layout. Font metrics, wrapping and terminal metrics can
differ. Unknown or active unsupported source still refuses; source bytes and
mutation safety remain unchanged. Shape autofit requires its separate opt-in.

Within that policy, properties that PowerPoint lays out but native v1 paint does
not model are validated, dropped from the projection, and disclosed per element
by `pptx.inherited-text-properties-omitted` (a sorted list such as `a:lnSpc`,
`a:spcBef`, `a:spcAft`, `a:buClr`, `a:buSzPct`, `a:tabLst`, `a:rPr@spc`,
`a:rPr@strike=noStrike`, `a:buFont@panose`, `a:endParaRPr`). Authored `a:br`
continues in a bullet-free paragraph at the same left margin, a paragraph
without runs becomes one blank space run carrying its end-mark metrics, and an
`a:t` with undeclared edge whitespace is projected as preserved; all three are
listed in the same disclosure. `a:buFontTx`/`a:buSzTx`/`a:buClrTx` cancel a
lower layer's bullet font/size/color before leaving the projection. Paint-active
properties (strikethrough,
underline, baseline shifts, highlights, automatic numbering, picture bullets)
still refuse. A shape `fontRef` without an authored color contributes only its
typeface. The `source-latin-inheritance-approximate-v1` identity names the
declared layer order and Latin-only profile, which are unchanged; the widened
property handling is disclosed per element by the omissions code rather than by
a new policy identity, because hosts and the preview server re-validate the v1
string as the inherited-text contract.

Without that opt-in, strict extraction still projects one narrower read-only
lane: when `p:presentation/p:defaultTextStyle` is present and a non-placeholder
shape's text is not self-contained, the presentation's explicit `lvlNpPr`
levels are merged beneath the shape-local `a:lstStyle` levels (local wins) and
the exact paragraph extractor runs on that owned projection. The element becomes
preserve-only and carries `pptx.presentation-text-style-preview`; mutation is
refused. Every layer is validated before precedence is applied: a `defPPr` in
either layer, any unmodeled level attribute or child (for example `defTabSz`,
`rtl`, `eaLnBrk`, `latinLnBrk`, `hangingPunct`, spacing), duplicate levels, and
malformed values refuse the text instead of being merged by guesswork. Text that
is already self-contained stays exact and editable; placeholders, whose chain
runs through master `txStyles`, and the inherited preview lanes are untouched.
This is a declared projection of explicit matching levels, not a qualification
of PowerPoint's inheritance precedence.

`NativePPTXExtractOptions.AllowSourceFrameAutoFitPreview` explicitly permits a
read-only preview of otherwise supported `spAutoFit` text in its saved source
frame. Its native `textBody.autoFit` is `shape-source-frame`, with a persistent
approximation diagnostic and non-editable status. No content-dependent resizing
is performed; frame size, text layout, and overflow or clipping may differ from
PowerPoint. Normal extraction and mutation remain strict. Malformed autofit,
unsupported vertical modes, unsupported fonts and other independent gaps remain
refusals.

The same option admits `a:normAutofit` as a read-only approximation of the
values PowerPoint itself authored: canonical `fontScale` (1%–100%, default 100%)
scales every resolved run size, rounded half up to whole hundredths of a point,
and canonical `lnSpcReduction` (0%–99%, default 0%) travels in the contract as
`textBody.lineSpacingReductionPercent1000` (thousandths of a percent, omitted
when zero) for the approximate renderer to reduce its line pitch by. The
frame is never resized, `textBody.autoFit` stays `none`, the element becomes
preserve-only, and `pptx.autofit-authored-scale-approximate` names both values.
Validation keeps the reduction bound to that disclosure: it may only appear on a
parsed, non-editable element that carries the warning, and never on a table cell.
Percent-string forms and out-of-range values refuse. `numCol`/`spcCol` (1–16
columns, nonnegative spacing) travel as `textBody.columnCount` /
`textBody.columnSpacingEmu` for the approximate renderer to flow the shaped lines
through, disclosed by `pptx.text-columns-approximate`. The fields are emitted
together, only when `numCol` is at least 2 and the saved frame still leaves a
positive width for every column; otherwise the same disclosure reports the
single-column fallback. `rtlCol` stays a refusal, so column order is always left
to right. Strict extraction keeps refusing `a:normAutofit` and multiple columns.

```bash
go get github.com/injectinglabs/injoffice/go/pptxpatch
```

```go
deck := pptxpatch.ExampleDeck()
bytes, err := pptxpatch.BuildPPTX(deck)
if err != nil {
	return err
}
```

The model supports text, basic shapes and connectors, tables, pictures, charts, transitions, and a bounded animation set. Unknown PPTX features are not promised to survive a parse-and-rebuild cycle; use surgical archive patching when preservation of unmodeled parts is required.

For the authoritative native Office boundary, use `ExtractNativePPTX` with a
trusted passthrough capability issuer. Documents containing DrawingML groups
require the issuer to implement `NativePassthroughTransactionalTokenFactory`, so
all capabilities publish only after one successful atomic commit:

```go
native, err := pptxpatch.ExtractNativePPTX(pptxBytes, pptxpatch.NativePPTXExtractOptions{
	Previous: previousNativeDeck,
	TokenFactory: tokenIssuer,
})
```

This path validates a bounded OPC graph, exact content and relationship types,
OOXML namespaces, source anchors, resource limits, and the final
`pptx-native/v1` contract. It never returns raw XML or package paths as
passthrough tokens. Unsupported fidelity is capability-backed and diagnostic, or
the extraction fails closed. Parsed picture assets are source-only: their source
part metadata never authorizes a read, and the trusted host must resolve the
asset passthrough capability against the source revision, actual media part,
digest, byte length, and issuance reason before serving bytes. See
[`docs/PPTX-NATIVE-CONTRACT.md`](../../docs/PPTX-NATIVE-CONTRACT.md) for the
supported extraction subset and canonical XML declaration rule. The current
editable AutoShape subset is renderer-free and exact: unadjusted rectangle,
ellipse, triangle, and diamond presets with explicit sRGB/no fill and explicit
solid outline metadata. Custom geometry, unresolved theme paint, gradients, patterns,
effects, unmodeled color transforms, and unsupported transforms become
capability-backed refused elements; they are not flattened to a nearby shape.
Untransformed theme scheme colors, documented `schemeClr` tint/shade/lumMod
transforms, and latin theme-font tokens are materialized into the existing
color and `fontFamily` fields when the relationship-routed theme and master
color map supply exact snapshots. Alpha-only and unmodeled color transforms
remain object-local refusals.

Static text previews accept the Boolean `dirty`, `smtClean`, and `err` run/default-run
checking flags. They are validated before style precedence and omitted only from
the owned paint projection; original package bytes remain untouched. Affected
text, shape, and table targets stay preserve-only because text replacement does
not round-trip those flags. This does not permit unsupported language/layout
properties, symbol-font bullet substitution, or shape/font autofit. Autofit and
unsupported inherited-property diagnostics name the remaining blocker.

Nonempty paragraphs may preserve an exact end mark containing only checking
flags and a language identical to the final visible run's resolved language.
End-mark font size, styling, different language, and empty/trailing-empty runs
remain unqualified. Accepted end marks stay read-only, with an explicit source
preservation diagnostic and a prewrite mutation refusal.

DrawingML `vert="vert"` projects read-only `writingMode="vertical-clockwise"`.
Per ECMA-376 Part 1 §20.1.10.83, lines rotate clockwise and progress leftward.
The renderer swaps the inner text frame's inline/block limits and transforms
only text, independently of shape/group rotations. The bounded preview accepts
non-bulleted ASCII Latin runs with LTR shaping; stacked, East Asian, RTL, other
vertical modes, and vertical table cells remain unsupported. Existing native
line-box policy labels still apply; this is not a claim of Office equivalence.

AutoShape frame rotations of 90, 180, and 270 degrees are projected as optional
`transform.quarterTurns` and rendered around the source frame center. Their
affine coefficients are exact integer values; 90/270-degree frames requiring
fractional-EMU centers remain refused. Normal horizontal shape text rotates
with its frame. Vertical modes outside the bounded clockwise profile, arbitrary-angle rotation,
flipped frames, autofit, and unresolved preset/theme paint remain unsupported.
Rotated source targets are preview-only and explicitly reject mutations; this
does not widen authoring or editing permissions.

The parsed connector subset projects unrotated `p:cxnSp` straight-line geometry
with positive X/Y extents and a complete explicit sRGB or documented theme
solid stroke. Named `headEnd`/`tailEnd` types, including `w`/`sz` values `sm`/`med`/`lg`
that map onto integer EMU as 2/3/5 × `stroke.widthEmu`, become exact
`headArrow`/`tailArrow` flags. Horizontal/vertical zero-extent
lines remain outside the v1 transform contract and fail extraction closed.
Unknown arrow types or sizes, unresolved theme paint, non-solid
dash, effects, and bent/custom geometry become exact capability-backed
refusals. Opaque charts stay preserve-only: an exact PNG/JPEG preview
relationship is painted at the frame EMU, and a missing preview remains a
RenderTree refusal rather than an invented chart renderer. Connector flip parity is retained for the straight line geometry
where endpoint reversal is visually equivalent. Grouped connectors retain
their authored order and inherit the exact DrawingML group affine without an
element-bounds clip that would truncate the centered stroke.

The parsed table subset projects only unrotated `p:graphicFrame` / `a:tbl`
grids whose frame extent exactly matches their positive explicit column and row
tracks. Every cell must be unmerged and self-contained: explicit sRGB/no fill,
four explicit no-fill borders, fixed horizontal text direction/top anchor,
explicit overflow and margins, and exact native paragraphs/runs. Table styles,
banding, inheritance or theme paint, effects, visible/partial borders, merges,
charts, and embedded objects remain exact capability-backed passthrough instead
of being approximated. Grouped tables retain source order and the exact parent
DrawingML affine.

Exact source-backed text and AutoShape property edits use
`ApplyNativePPTXMutations`. Every batch must include the `sourceRevision` from
the matching extraction and every operation must bind the target element's
source fingerprint. The applier edits only bounded XML spans in the original
slide part, raw-copies and verifies every untouched OPC entry, then reopens and
extracts the result to validate the requested values. It refuses structural
edits, non-self-contained text, inherited or unsupported shape semantics, and
stale revisions rather than rebuilding a presentation from the render model.
Exact text and AutoShape descendants of projected DrawingML groups are valid
targets; their ancestor group transforms, child coordinate spaces, order,
compatibility, and passthrough inventory remain unchanged. Text replacement
retains the source `a:bodyPr` layout and refuses literal tabs or line breaks.
The rewritten ZIP also preserves and verifies the original archive EOCD
comment in addition to raw-copying every untouched entry.

When the request comes from the shared Office mutation envelope, use
`ApplyNativePPTXMutationPayload`. It strictly decodes at most 3 MiB of native
JSON, rejects duplicate or unknown fields and trailing JSON, independently
checks the envelope's `sha256:<digest>` exact-byte CAS, and requires the native
payload's `expectedSourceRevision` to be the corresponding `rev-<digest>`.
Successful edits additionally require an actual semantic OOXML change and
preserve the presentation/slide/element topology plus the exact bidirectional
unsupported diagnostic and passthrough inventory.

### Read-only placeholder inheritance

Top-level title/body placeholders resolve one unique relationship-bound layout
index and master type. Complete geometry, body properties, and explicit list
levels cascade into a read-only projection; original slide runs and source
fingerprints are retained. Ancestor authoring prompt text is never copied into
the slide. Explicit level 0–8 prompt paragraph properties and `defRPr` defaults
can supply missing or equal list-level properties. Duplicate levels, conflicting
defaults, implicit styled levels, and run/end-mark styling refuse; no competing
source is guessed. Local slide paragraph/run overrides still win, including
explicit nested list levels. Competing placeholder matches, grouped placeholders,
and unmodeled master artwork remain outside this subset. This is not whole-master
or PowerPoint fidelity.

When that exact chain does not qualify and `AllowInheritedTextPreview` is set,
`title`, `ctrTitle`, `subTitle`, `body`, `obj` and untyped placeholders resolve
through a declared approximation instead: layout match by `idx` (falling back to
an exact `type` match), master match by title/body family, and the cascade
presentation `defaultTextStyle` → master `titleStyle`/`bodyStyle` → master placeholder `lstStyle` → layout
placeholder `lstStyle` → slide `lstStyle` → local properties, fed through the
inherited text preview above. Master/layout/slide `bodyPr` merge with the
autofit child as one replaceable slot; the merged result is validated by the
text-body layout policy. Solid frame paint cascades the same way, nearest layer
wins: an exact sRGB or documented theme `a:solidFill` and an `a:ln` solid fill
(its declared `w`, or a 9525 EMU hairline when none is declared) are painted
read-only, and the element becomes a `rect` preset shape carrying the existing
AutoShape `fill`/`stroke` fields; a nearer `a:noFill` paints nothing. Gradient,
pattern and picture fills, effects, shape styles, dashes, compound lines, caps
and joins are not painted and are disclosed on ancestors; on the slide
placeholder itself they refuse. Ancestor prompt paragraphs are never painted;
the element carries `pptx.placeholder-inheritance-approximate` describing the
inherited frame and listing what was not painted, plus the inherited-text
disclosures. A slide placeholder without a text body keeps its inherited frame
and paints no text (PowerPoint shows no prompt); because PowerPoint slideshow
and export hide placeholders without text, a painted frame on an empty
placeholder is disclosed as an approximation. `hidden` placeholders refuse, as
do date, footer, slide-number, picture, chart, table, media and diagram
placeholders. `element.placeholder` reports the family (`obj` and untyped report
`body`).
