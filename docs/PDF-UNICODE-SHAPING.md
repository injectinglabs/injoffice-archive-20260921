# Embedded Unicode form appearances: LTR clusters

Caller-supplied `textAppearance: { fontBytes }` now shapes supported single-line LTR text with HarfBuzz. Covered Latin, Greek, Cyrillic and the existing East Asian repertoire retain exact source values; combining marks, composition and ordinary ligatures are supported. This is the first cluster-aware milestone. Cross-script-specific combining marks, bidi paragraph resolution, Arabic/Hebrew and other contextual script itemization remain subsequent work, not completed Unicode coverage.

The caller supplies a standalone fixed TrueType outline font. The complete face is embedded, bounded to 16 MiB, so saving a short value may produce a larger PDF than the earlier subset profile. Font discovery, fallback, collections, CFF, variable/color font formats and vertical/multiline layout are not introduced here. Source PDF/font bytes remain unchanged. Existing ownership, widget rotation, plain-field, mixed standard-choice and failure-atomicity rules remain in force.

## Glyphs, source text and extraction

HarfBuzz returns original UTF-16 cluster offsets separately from glyph IDs and x/y placements/advances. A source-semantic CID is allocated for each glyph/text mapping, with explicit CIDToGIDMap and ToUnicode. Two source strings sharing an outline cannot corrupt each other's extraction mapping. Maps are staged until the form value is accepted.

A cluster that naturally partitions into its original scalar glyphs maps each scalar separately. A composed or ligature glyph maps its complete original cluster. If a cluster produces additional glyphs that cannot receive a nonempty exact source span, those continuation glyphs are painted from the font's own outlines at the shaped positions. Quadratic curves are converted exactly to cubic curves; winding and holes are retained. These continuation outlines do not receive TrueType hinting, which can cause small raster differences at low resolution. The main glyph remains embedded native text. No invented character, empty ToUnicode destination, normalization or transliteration is used.

Positioned runs carry an ActualText span as additional replacement semantics. Saved/flattened appearance tests use PDF.js `getTextContent({ disableNormalization: true })` to verify decomposed text, ligatures, reordered stacked marks and supplementary scalars. Poppler proof covers the same representation. PDF.js currently ignores ActualText in ordinary text extraction, so ActualText alone is insufficient: an unmapped continuation CID becomes an unrelated fallback character. The outline policy avoids that contamination. Reader whitespace/normalization heuristics remain distinct from canonical form `/V`, which preserves the exact input including edge spaces.

## Layout and bounds

The default appearance provider retains authored colors, borders, alignment, rotation and field behavior. A local opaque layout token measures the entire shaped run, including cross-space kerning, without storing the token. Exactly one expected text operator is replaced. Arbitrary offsets use independent glyph matrices; outlines leave the text object and inherit its monochrome fill.

Auto-sized mark-bearing text measures the union of positioned glyph ink and logical advances, including negative bearings and tall stacked marks. Fixed-size fields keep the existing clipping behavior. Limits are 4096 UTF-16 units per value, 16384 shaped glyphs, 65535 semantic CIDs per operation and 65536 continuation path commands per run. An unpartitionable source cluster may contain at most 256 UTF-16 units because a ToUnicode destination is limited to 512 bytes. Invalid metrics, paths, missing glyphs or unsupported scripts produce explicit refusal rather than partial text.

## Primary references

- [HarfBuzz clusters](https://harfbuzz.github.io/clusters.html): glyph count and source character count are independent; cluster offsets retain their relationship.
- [OpenType GPOS](https://learn.microsoft.com/en-us/typography/opentype/spec/gpos): placement and advance have independent horizontal and vertical components.
- [PDF 32000-1:2008](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/PDF32000_2008.pdf), §§9.7, 9.10.3 and 14.9.4: composite fonts, ToUnicode destination strings and replacement text.
- [Unicode Bidirectional Algorithm](https://www.unicode.org/reports/tr9/): required paragraph resolution/itemization for the next milestone; reversing a string is not a bidi implementation.
