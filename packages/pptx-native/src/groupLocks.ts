/** Diagnostic code the Go extractor attaches to a group it kept although its
 * `p:cNvGrpSpPr` carried the `a:grpSpLocks` PowerPoint writes whenever shapes
 * are grouped. Every attribute that element can carry is an editing lock the
 * authoring UI enforces, so none of them can move a pixel -- but none of them
 * travels on the wire either, so such an element must stay `preserveOnly`,
 * carry a parsed source anchor and never be mutated. */
export const PPTX_GROUP_LOCKS_PRESERVED_CODE = 'pptx.group-locks-preserved'
