# pptxpatch

`pptxpatch` reads and writes a bounded, editable PowerPoint OOXML model.

`NativePPTXExtractOptions.AllowSourceFrameAutoFitPreview` explicitly permits a
read-only preview of otherwise supported `spAutoFit` text in its saved source
frame. Its native `textBody.autoFit` is `shape-source-frame`, with a persistent
approximation diagnostic and non-editable status. No content-dependent resizing
is performed; frame size, text layout, and overflow or clipping may differ from
PowerPoint. Normal extraction and mutation remain strict. Malformed autofit,
font scaling, unsupported vertical modes, unsupported fonts and other independent gaps remain
refusals.

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
