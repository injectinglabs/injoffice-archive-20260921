# Embedded Unicode form appearances: clusters and bidi

Caller-supplied `textAppearance: { fontBytes }` shapes supported single-line mixed-direction text with HarfBuzz. Covered Latin, Greek, Cyrillic, Arabic, Hebrew and the existing East Asian repertoire retain exact source values; combining marks, composition and ordinary ligatures are supported. Unicode 17 paragraph resolution handles levels, brackets and isolates before directional/script itemization. HarfBuzz shapes each item with full original source context. Optional `direction` (`auto`, `ltr`, `rtl`; default `auto`) and `language` (default `und`) select paragraph direction and shaping language. Cross-script-specific combining marks and other contextual script itemization remain subsequent work, not completed Unicode coverage.

The caller supplies a standalone fixed TrueType outline font. The complete face is embedded, bounded to 16 MiB, so saving a short value may produce a larger PDF than the earlier subset profile. Font discovery, fallback, collections, CFF, variable/color font formats and vertical/multiline layout are not introduced here. Source PDF/font bytes remain unchanged. Existing ownership, widget rotation, plain-field, mixed standard-choice and failure-atomicity rules remain in force.

## Glyphs, source text and extraction

HarfBuzz returns original UTF-16 cluster offsets separately from glyph IDs and x/y placements/advances. A source-semantic CID is allocated for each glyph/text mapping, with explicit CIDToGIDMap and ToUnicode. Two source strings sharing an outline cannot corrupt each other's extraction mapping. Maps are staged until the form value is accepted.

A cluster that naturally partitions into its original scalar glyphs maps each scalar separately. A composed or ligature glyph maps its complete original cluster. If a cluster produces additional glyphs that cannot receive a nonempty exact source span, those continuation glyphs are painted from the font's own outlines at the shaped positions. Quadratic curves are converted exactly to cubic curves; winding and holes are retained. These continuation outlines do not receive TrueType hinting, which can cause small raster differences at low resolution. The main glyph remains embedded native text. No invented character, empty ToUnicode destination, normalization or transliteration is used.

Positioned runs carry an ActualText span as additional replacement semantics. Saved/flattened appearance tests use PDF.js `getTextContent({ disableNormalization: true })` to verify decomposed text, ligatures, reordered stacked marks and supplementary scalars. Poppler proof covers the same representation. PDF.js currently ignores ActualText in ordinary text extraction, so ActualText alone is insufficient: an unmapped continuation CID becomes an unrelated fallback character. The outline policy avoids that contamination. Reader whitespace/normalization heuristics remain distinct from canonical form `/V`, which preserves the exact input including edge spaces.

For RTL, generic reader extraction is not universally exact. With DejaVu Sans and source `السلام عليكم`, PDF.js with normalization disabled extracts `السالم عليكم`: its bidi heuristics reverse the internal lam-alef span despite canonical ToUnicode `<06440627>`. ActualText and form `/V` retain the exact logical source; mappings are never reversed to compensate for a reader. Mixed-direction spacing and isolates can also change in generic extraction. Use the form value API when exact field source is required.

## Unicode data and attribution

The PDF-local bidi algorithm derives from MIT-licensed bidi-js 1.1.0 (Jason Johnston). It receives Unicode scalars instead of UTF-16 code units. Generated Unicode 17 data replaces the upstream Unicode 13 tables. The local reorder helper corrects trailing whitespace indexing for nonzero line starts and omits source string rewriting/mirroring; HarfBuzz paints directional glyphs. See `packages/pdf/third-party/BIDI-JS-LICENSE.txt` and root `LICENSE-UNICODE.txt`, both included in package legal files.

`scripts/generate-pdf-unicode-bidi.mjs --check` regenerates tables offline from SHA-256-pinned official inputs in `packages/pdf/testdata/unicode17`. The generated module retains the hashes. Tests run all 91,707 BidiCharacterTest cases and 770,241 BidiTest paragraph permutations through rule L2. Glyph attachment and mirroring are separately tested at the shaping/PDF boundary; passing the scalar suites alone is not a claim about every font or script. Shared Unicode tables remain unchanged.

## Layout and bounds

The default appearance provider retains authored colors, borders, alignment, rotation and field behavior. A local opaque layout token measures the entire shaped run, including cross-space kerning, without storing the token. Exactly one expected text operator is replaced. Arbitrary offsets use independent glyph matrices; outlines leave the text object and inherit its monochrome fill.

Auto-sized RTL and mark-bearing text measures the union of positioned glyph ink and logical advances, including negative bearings and tall stacked marks. Fixed-size fields keep the existing clipping behavior. Limits are 4096 UTF-16 units per value, 16384 shaped glyphs, 65535 semantic CIDs per operation and 65536 continuation path commands per run. An unpartitionable source cluster may contain at most 256 UTF-16 units because a ToUnicode destination is limited to 512 bytes. Invalid metrics, paths, missing glyphs or unsupported scripts produce explicit refusal rather than partial text.

## Primary references

- [HarfBuzz clusters](https://harfbuzz.github.io/clusters.html): glyph count and source character count are independent; cluster offsets retain their relationship.
- [OpenType GPOS](https://learn.microsoft.com/en-us/typography/opentype/spec/gpos): placement and advance have independent horizontal and vertical components.
- [PDF 32000-1:2008](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/PDF32000_2008.pdf), §§9.7, 9.10.3 and 14.9.4: composite fonts, ToUnicode destination strings and replacement text.
- [Unicode Bidirectional Algorithm](https://www.unicode.org/reports/tr9/): Unicode 17 paragraph resolution, bracket pairing, isolates and visual reordering.
