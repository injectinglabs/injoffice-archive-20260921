# docxpatch

`docxpatch` surgically edits WordprocessingML paragraphs and selected document features without regenerating unrelated DOCX parts.

```bash
go get github.com/injectinglabs/injoffice/go/docxpatch
```

```go
paragraphs, err := docxpatch.Extract(input)
if err != nil {
	return err
}
output, err := docxpatch.Apply(input, []docxpatch.Edit{{
	Op:    "set",
	Index: paragraphs[0].Index,
	Text:  "Updated heading",
}})
```

Paragraphs with unsupported non-text content are refused rather than flattened. Additional APIs cover images, charts, notes, themes, and text watermarks; consult `go doc` for their explicit scope.

## Native v1 extraction

`ExtractNativeDocumentV1` is the HTML-free, fail-closed DOCX import boundary.
It resolves the root `officeDocument` relationship instead of assuming
`word/document.xml`, supports Transitional and Strict WordprocessingML, and
returns the versioned `NativeDocumentV1` contract with exact part-byte anchors.

```go
native, err := docxpatch.ExtractNativeDocumentV1(input)
if err != nil {
	return err
}
wireJSON, err := docxpatch.EncodeNativeDocumentV1(native)
```

For a later source revision, pass the previously issued contract so unchanged
objects retain their InjOffice identities even if sibling paths move:

```go
next, err := docxpatch.ExtractNativeDocumentV1WithOptions(updatedInput,
	docxpatch.NativeExtractionOptions{Previous: native})
```

An application can set `DocumentID` on first import. Without one, the initial
document identity is derived from the main XML only, so changing unrelated
media or metadata does not churn it. `Previous` must be a valid v1 document
for the same case-equivalent OPC main part. Unique eight-digit Office 2010
`w14:paraId` values in the positive 31-bit range seed paragraph identity;
invalid or duplicate values stay preserved in the source anchor but are not
trusted as identifiers.

The current native subset includes body/header/footer/note/comment paragraphs,
Unicode text, basic controls and native references, conservative paragraph/run
properties, basic tables, sections and page geometry, and single DrawingML
pictures with direct image relationships. Images, relationship parts, styles,
numbering, themes, font tables, and every other non-story OPC part remain
fingerprinted `preserve-verbatim` inventory entries even when the read-only
resolver below consumes them.

Footnote and endnote stories retain the exact main-part relationship ID and an
explicit content, separator, or continuation-separator role. Content stories
accept the matching in-story `w:footnoteRef`/`w:endnoteRef` label as a typed
self-reference. The special separator IDs are exact (`-1` and `0`), ordinary
IDs must be unique, and wrong or ambiguous relationships/content types fail
the whole extraction. This is structural extraction only: qualified native
pagination supports paintable paragraph content and exact conventional
`w:separator`/`w:continuationSeparator` instruction leaves in their matching
sentinels. Custom numbering, nested tables, drawings, fields, misplaced
instructions, and other unmodeled note markup remain fail-closed.
The `0` continuation sentinel is dormant during fitting pagination; its
preserved contents cannot refuse until continuation is actually required.

Fields, hyperlink semantics, tracked/wrapped content, nested tables, VML and
objects, crop/effect-bearing or ambiguous pictures, custom table/section
geometry, and unknown extension markup produce explicit unsupported records.
Their owning paragraph/table/drawing is read-only. Extraction does not claim a
lossy whole-document regeneration. Safe paragraphs advertise `text.replace`;
`ApplyNativeTextMutationsV1` consumes exact paragraph or text-run IDs plus XML
fingerprints, uses the full `source.package_sha256` value for exact-byte CAS,
rejects stale revisions and unsupported markup, and returns a reopened native
contract with part-hash preservation evidence. The shorter opaque `revision`
field is never save authority. `DecodeNativeDOCXTextMutationPayloadV1`
strictly bounds the format payload and rejects duplicate, unknown, or trailing
JSON before mutation. Envelope handlers should use
`ApplyNativeTextMutationPayloadV1`, which proves the outer exact-package CAS
before decoding or applying the payload. Semantic no-ops, changed unsupported inventories, new or
lost targets, and document-topology drift are refused. Paragraph-level
replacement is deliberately limited to paragraphs containing exactly one
modeled text node; callers can target individual text runs in other safe
paragraphs.

Package parsing is bounded to 128 MiB compressed, 256 MiB total expanded,
64 MiB per part, 16 MiB per XML part, 10,000 entries, a 200:1 per-entry
compression ratio (with 1 MiB slack), 100,000 XML elements, and 64 XML levels.
Duplicate/case- or percent-equivalent parts, unsafe or missing internal
relationship targets, namespace spoofing, DTD/directive/entity tricks, and
ambiguous content-type/relationship metadata are rejected.

## Native style and numbering resolution

`ResolveNativeDocumentLayoutV1` builds a separate, deterministic
`injoffice.docx.resolved-layout` v1 projection on top of native extraction. It
is keyed by the durable native paragraph/run/table IDs and does not alter or
extend the persisted `NativeDocumentV1` wire contract.

```go
layoutInput, err := docxpatch.ResolveNativeDocumentLayoutV1(input)
if err != nil {
	return err
}
layoutJSON, err := docxpatch.EncodeNativeResolvedLayoutInputV1(layoutInput)
```

The HTML/DOM-free resolver discovers styles, numbering, themes, and font tables
through the main part's Strict or Transitional relationships at arbitrary
case-equivalent OPC locations. It resolves document defaults, base-to-derived
paragraph/character style chains, style toggles, direct formatting, logical
paragraph direction/indents, and ordinary numbering levels/overrides. Direct
paragraph properties come from the exact anchored story XML so spacing and
indentation are not limited by the persisted v1 subset. Direct pagination-sensitive
`pPr`, spacing, indent, keep, page-break, widow, bidi, and line-break/control
nodes must have exact attributes, children, and XML-whitespace-only text at
extraction; smuggled direct structure is diagnosed. Every paragraph also
exports its resolved paragraph-mark run properties (document defaults,
paragraph-style run properties, and direct `pPr/rPr`) so a native line engine
can measure blank, hidden-only, and control-only paragraphs without inventing
text or a platform font default.

Extraction admits direct paragraph-mark Latin font names, size, bold/italic,
RTL/hidden state, RGB color, and language when their source structure is exact.
Paragraphs with these properties remain preservation-only for editing. Duplicate
properties, character-style references, unknown attributes/children, and other
mark effects still refuse native layout rather than inheriting guessed values.

The output exposes explicit Latin fonts, half-point sizes, RGB colors,
language, bidi/RTL/hidden state, supported highlights/underlines, paragraph
alignment/spacing/indent/keep/page-break/widow state, and ordinary decimal,
letter, Roman, and bullet list markers. Exact `w:themeColor` / `w:themeFill`
values resolve through leaf `a:srgbClr` or `a:sysClr lastClr` theme slots into
those RGB fields. Exact `asciiTheme`/`hAnsiTheme` latin slots resolve to the
theme typeface name. Tint/shade, automatic color, complex-script/
East-Asia font variants, automatic or character-unit spacing, picture bullets,
and conditional table-style region effects remain preserve-verbatim with scoped
diagnostics. Simple table styles may project whole-table borders and clear fills.
Malformed, ambiguous, namespace-spoofed, wrong-content-type, cyclic, or
resource-hostile inputs are rejected or conservatively excluded from the
cascade.

This model is input to a future font loader, shaper, line/table layouter, and
paginator. It does not shape glyphs, break lines, calculate pages, paint a
canvas, mutate anchors, or regenerate OOXML.

## Canonical DOCX font inventory and exact resolver

`ExtractNativeDOCXFontInventoryV1` emits a separate
`injoffice.docx.font-inventory` contract. It binds the durable document ID,
opaque document revision, full package SHA-256, actual main part and hash, the
main-to-font-table relationship (including lexical target) and owning relationship-part hash, the font
table hash, and the font-table relationship-part hash. Every supported face
also binds its authored family/alias and regular/bold/italic slot, exact
relationship ID/type/lexical target, actual-case target part and content type, stored length
and digest, deobfuscated content digest, font key, subset flag, and verified
OS/2 embedding rights. Inventory and native-text-manifest digests cover their
complete versioned canonical JSON projections across Go and Node.

V1 deliberately accepts only internal, dialect-correct OOXML font
relationships to `obfuscatedFont` parts with a canonical nonzero font key and
a bounded standalone sfnt envelope. Restricted, bitmap-only, reserved or
conflicting licensing flags, malformed sfnt directories, missing OS/2 tables,
and subset/no-subsetting conflicts are refused. The emitted native-text
manifest contains only `document` sources with mandatory decoded digests and
an empty fallback list; it never invents PostScript names, weights, oblique
faces, host paths, system fonts, or substitutions.

```go
inventory, err := docxpatch.ExtractNativeDOCXFontInventoryV1(input)
if err != nil {
	return err
}
asset, err := docxpatch.ResolveNativeDOCXFontAssetV1(input, inventory,
	docxpatch.NativeDOCXFontResolutionRequestV1{
		Family: "Example Sans", Weight: 400, Style: "normal", Stretch: 100,
})
```

Persistent Node page-paint bridges should pass
`EncodeNativeDOCXFontInventoryV1(inventory)` as `font_inventory_json` and use
`ResolveNativeDOCXPagePaintFontAssetsV1(input, inventory)` for the exact
`font_assets` array. The helper reconstructs the inventory from the package
before emitting face slot, resource ID, content digest, explicit collection
index, and canonical base64 bytes. A bridge must not accept a caller-provided
font manifest or substitute any of those identities.

The resolver reopens the package and reconstructs the complete inventory
before returning decoded bytes. A part path or digest supplied by a caller is
never authority. The bounded sfnt work here stops at the table directory and
OS/2 `fsType`; glyph outlines, cmap, names, metrics, shaping, outline
serialization, and transport remain provider responsibilities. Font-table and
font parts stay in the structural extractor's `preserve-verbatim` inventory,
so this read-only projection does not weaken surgical package preservation.

## Native pagination settings attestation

`ExtractNativePaginationSettingsV1` emits the separate versioned
`injoffice.docx.pagination-settings` v1 projection consumed by the native
shaper/paginator. It relationship-discovers `settings.xml` from the actual main
part and binds the package SHA-256, actual owning `.rels` part and SHA-256,
relationship ID, canonical settings part name, and settings SHA-256 to the same
native document identity and revision.

```go
settings, err := docxpatch.ExtractNativePaginationSettingsV1(input)
if err != nil {
	return err
}
settingsJSON, err := docxpatch.EncodeNativePaginationSettingsV1(settings)
```

An absent part attests Word's 720-twip default tab stop but also implies the
schema-default compatibility mode 12, so pagination v1 refuses it. The only
supported modern profile requires explicit compatibility mode 15 and accepts positive tab stops,
`doNotCompress` character spacing, the even/odd header-selection flag, and a
bounded set of settings proven neutral for the paginator's already narrow body
subset. Mirror margins, top gutter, legacy/unknown compatibility options, and
unknown layout-affecting settings remain preserved but produce an
`unsupported` profile with exact part/path diagnostics. Accepted elements are
structurally exact: unexpected attributes, children, non-whitespace text, and
duplicate singletons refuse the profile. Theme-font language, shape defaults,
attached templates, native math defaults, placeholder/revision display
settings, enabled field-result updates, and other settings not proven
irrelevant to the joined resolved/shaped inputs are not labeled neutral.
`updateFields=false` is attested explicitly. Content-type identity uses
ASCII-only case folding. This projection does not weaken or add fields to
persisted `NativeDocumentV1`. These joins are structural provenance, not an
authorization or signature boundary; hosts must retain the trusted extractor
result with the package it describes.
