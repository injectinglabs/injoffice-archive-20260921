# Native PPTX preview worker

The optional Boolean request `inherited_text_preview` permits explicitly marked
read-only inherited text approximations. Positive responses include
`inherited_text_preview_count` and the exact `inherited_text_policy` value
`source-latin-inheritance-approximate-v1`; absent markers omit both. The HTTP
preview selects this separately with `text=source-inherited` and validates those
fields against the extracted source, revision and slide inventory. It does not
implicitly enable `autofit=source-frame`. Callers must show the persistent
approximation warning, not report complete native fidelity.

Private Node 22+ adapter for real PPTX uploads. The Go helper extracts the original
package and invokes this worker through bounded, length-prefixed JSON. No hosted
service is required. Enable the helper with `-pptx-preview-worker` pointing to
`dist/worker.js` and `-pptx-font-manifest` pointing to an absolute local JSON file:

```json
{"version":1,"faces":[{"family":"DejaVu Sans","weight":400,"style":"normal","path":"/absolute/DejaVuSans.ttf","sha256":"sha256:<64 lowercase hex characters>"}]}
```

The operator is responsible for supplying appropriately licensed fonts and exact
family/style mappings. Each file is digest-checked and measured by the pinned
HarfBuzz implementation. No aliases, platform-font discovery, or silent fallback
are used. Missing family/weight/style combinations refuse visibly. Limits are
32 faces, 16 MiB per font, 64 MiB total font bytes, and 16 MiB framed JSON.

Before full extraction, the HTTP helper checks ZIP metadata: at most 8 MiB
uploaded, 2,048 entries, 8 MiB expanded per entry and 32 MiB expanded in total.
The entry count is checked after Go parses metadata from the bounded 8 MiB
upload; it is not a strict metadata-allocation cap. Safe directory records do
not participate in OPC part alias checks, but their declared sizes still count
toward admission budgets.
Duplicate/case/percent aliases, unsafe names, encryption, unsupported compression
and invalid local-entry bounds refuse before decompression. The extractor still
checks actual decoded sizes, CRCs, XML and package relationships. These limits
bound admitted package data, not total process memory. Cancellation is checked
around metadata traversal and synchronous extraction; the 45-second request
context cannot interrupt a synchronous extractor call. The Node subprocess has
its separate 30-second timeout and a V8 heap limit, not an OS memory sandbox.

The playground's native PPTX workbench asks for explicit upload consent and
verifies the response's package SHA-256, slide index, and slide count before
mounting bounded SVG paths. Source files are never rewritten by this path.
Measured mixed-run line boxes and anchors use `max-run-natural-v1`, not an
Office-equivalence claim. The existing approximate file preview stays available.
Embedded static PNG and baseline JFIF images are source-part/digest-bound and
decoded before display. DrawingML positive source-edge crops are replayed without
rewriting bytes; malformed rasters and browser decode failures clear native
success. Browser limits are 16 million pixels per image and 32 million total.
Typed source `headEnd`/`tailEnd` descriptors replay triangle, open arrow, stealth,
diamond, and oval endpoints under the explicit `arrow-v1` preview policy. Named
small/medium/large dimensions use 2/3/5 times the stroke width; omitted dimensions
use medium. Closed shapes are filled, open arrows are stroked; endpoints point
outward from their source straight line. This deterministic geometry is not an
Office-equivalence claim. Legacy boolean-only endpoints stay visibly unqualified.
Source names follow Microsoft's [DrawingML head-end documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.headend?view=openxml-3.0.1):
`type`, `w`, and `len` (not the nonstandard `sz`). Original package bytes remain
unchanged, including omitted attributes; the transport keeps omissions explicit.

Run `node scripts/smoke-pptx-native-preview-browser.mjs` from the built workspace
for real-file upload, HarfBuzz glyph, anchor, hanging bullet, cropped-quadrant
pixel, image failure, consent, and missing-font checks.
