# Embedded PDF text kerning

The existing `textAppearance: { fontBytes }` form option saves horizontal kerning from a caller-supplied fixed TrueType font. The same positioned advances determine left/center/right alignment and automatic font size. The playground's local font picker uses this behavior without an additional mode.

This remains a bounded single-line profile: existing script, mark, missing-glyph, glyph-alias, font-format and widget-ownership guards apply. Only nonnegative finite horizontal advances are accepted; horizontal/vertical offsets, vertical advances, substitution, combining marks and bidi/contextual layout remain unsupported. Choice appearances retain their independent standard-font/ASCII/12-point fit policy. Explicitly fixed text sizes can still clip.

## Saved operators and extraction

OpenType pair positioning may adjust either placement or advance, in either axis. This profile accepts only horizontal advance changes. It checks that the positioned glyph sequence matches the kerning-disabled subset encoder's sequence and retains the one-scalar-per-glyph Unicode mapping guard. See [Microsoft's GPOS pair-adjustment specification](https://learn.microsoft.com/en-us/typography/opentype/spec/gpos).

The saved font widths remain nominal. A PDF `TJ` array compensates between glyphs by `(nominalAdvance - positionedAdvance) * 1000 / unitsPerEm`. Positive PDF adjustments subtract horizontal distance; the source scalar encoding and ToUnicode mapping stay unchanged. Runs needing no adjustment retain `Tj`. See Adobe's [PDF Reference, Table 5.6, printed page311](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.4.pdf).

The implementation uses the existing appearance provider for style, clipping, rotation and default appearance handling. Its default autosizing splits words, so a local field/font facade supplies one layout token with the complete run width and the exact original encoding. The token never enters `/V`, an encoded text string or font resources. The adapter requires exactly one expected text-show operator before substituting `TJ`; unsupported provider behavior rejects the whole operation. It does not replace methods on the actual field/font instances.

Tests pin DejaVuSans2.37 `AV` nominal widths1401+1401 and pair adjustment−131 at2048 units/em. The saved correction is63.96484375 PDF text units. Tests also cover `To`, repeated/leading/trailing spaces, empty values, alignment, autosizing, rotation, invalid positioning, original Unicode extraction and mixed choice appearances. The browser smoke checks an actual font upload/apply/download and saved spacing operators. Renderer proof with a generated fixture demonstrates this implementation profile; it is not a general typography-fidelity claim.
