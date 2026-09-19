/** Diagnostic code the Go extractor attaches to a table whose cell paint was
 * resolved from the bounded built-in table style catalog. Such an element is
 * an explicitly labeled read-only preview: it must stay `preserveOnly`, carry
 * a parsed source anchor and never be mutated. */
export const PPTX_TABLE_BUILTIN_STYLE_PREVIEW_CODE = 'pptx.table-builtin-style-preview'

/** Declared paint policy named inside the preview diagnostic message. Fills
 * come from ECMA-376 table-part definitions of the catalog style with tints
 * mixed in linear sRGB and standard sRGB encoding; borders are uniform 1pt. */
export const PPTX_TABLE_BUILTIN_STYLE_POLICY = 'builtin-table-style-catalog-linear-srgb-tint-v1'

/** Diagnostic code the Go extractor attaches to a table it kept although its
 * p:cNvGraphicFramePr or p:nvPr carried metadata the exact projection does not
 * model: the `a:graphicFrameLocks noGrp="1"` PowerPoint writes on every table,
 * and the `p14:modId` extension it writes on every table it has edited.
 * Neither travels on the wire, so such an element must stay `preserveOnly`,
 * carry a parsed source anchor and never be mutated. */
export const PPTX_TABLE_NONVISUAL_PRESERVED_CODE = 'pptx.table-nonvisual-preserved'
