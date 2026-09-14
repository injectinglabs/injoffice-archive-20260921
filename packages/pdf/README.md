# @injoffice/pdf

Composable PDF inspection, page operations, text and image editing, annotations, forms, OCR layers, and whole-object redaction.

```bash
npm install @injoffice/pdf
```

```ts
import { applyPageOps, readInfo } from '@injoffice/pdf/browser'

const source = new Uint8Array(await file.arrayBuffer())
const info = await readInfo(source)
const rotated = await applyPageOps(source, [
  { type: 'rotate', pages: [info.pageCount], degrees: 90 },
])
```

Browser viewers must point pdf.js at a worker asset emitted by the host bundler. With Vite:

```ts
import workerURL from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { configurePdfWorker, PdfViewerDocument } from '@injoffice/pdf/browser'

configurePdfWorker(workerURL)
const viewer = await PdfViewerDocument.load(source)
```

For PDFs needing predefined character maps, standard fonts or image/color
decoders, serve the `cmaps/`, `standard_fonts/` and `wasm/` directories from the
**same installed `pdfjs-dist` version** as the worker. Preserve filenames and
included license files. Pass host-controlled directory URLs (with trailing
slashes) per document:

```ts
const viewer = await PdfViewerDocument.load(source, {
  cMapUrl: '/my-app/pdf-assets/cmaps/',
  standardFontDataUrl: '/my-app/pdf-assets/standard_fonts/',
  wasmUrl: '/my-app/pdf-assets/wasm/',
  useSystemFonts: false,
})
```

Packed CMaps are used. Resources load on demand; this does not upload the PDF.
No resource host or CDN is selected by the library. Supply trusted URLs subject
to your application's origin/CSP policy, not paths from document content.
Omitting options preserves PDF.js defaults, including its environment-dependent
system-font substitution policy. Supplying resources improves availability but
does not establish exact rendering or fix upstream Type3 differences. The
playground emits matching assets under its deployment base and disables system
font substitution for its viewer.

`renderPageToCanvas(viewer, page, canvas, zoom, pixelRatio, { signal, maxPixels })`
accepts an `AbortSignal` to cancel obsolete PDF.js render tasks. By default the
backing store is limited to 16 million pixels and 16,384 pixels per dimension;
large/high-DPI pages retain their CSS zoom but rasterize at reduced resolution.
These limits bound the canvas, not PDF parsing or total document memory.

`viewer.search(query, { signal })` checks cancellation between page extractions.
It reuses a bounded text cache (256 pages / 2 million characters); it does not
interrupt an already-running text extraction. `viewer.getOutlineTree()` retains
nested bookmarks, while `getOutline()` remains the flat compatibility API.
Call `viewer.destroy()` when replacing or closing a document.

`@injoffice/pdf/browser` contains the browser-safe viewer and page operations. The root entry additionally exposes Node.js PDFium, system-font, image, OCR, and redaction paths. Treat transformed files as untrusted input, apply resource limits in the host, and use the package's verification results for destructive redaction workflows.

### Explicit text form appearances

`applyFormValues(source, values)` retains its existing behavior: text values,
including Unicode, are written without replacing their appearance streams and
request viewer regeneration. To explicitly replace supported text appearances:

```ts
const result = await applyFormValues(source, [
  { name: 'customer', kind: 'text', value: 'Alex Smith' },
], { textAppearance: { font: 'Helvetica' } })
```

Choice values select the authored export value, which can differ from the
display label. The source options and field flags are preserved. Unknown values
are skipped unless the source dropdown already permits free-text editing;
ambiguous duplicate export values are skipped. Choice appearances still require
viewer regeneration. XFA forms refuse every form-update batch, including default
radio/choice edits, before pdf-lib can remove their source XFA data.

To save replacement choice artwork, separately request
`choiceAppearance: { font: 'Helvetica' }` (also `Times-Roman` or `Courier`).
This policy uses 12-point text and pdf-lib's choice appearance providers. It
supports noneditable dropdowns and single-selection lists with 1–64 unique,
nonempty printable ASCII display labels, at most 256 characters each. Export
values remain unchanged even when they differ from the painted labels. All
labels and list rows, including final descenders, must fit every owned widget
at 12 points. Widget rotations must be exactly 0, 90, 180 or 270 degrees. Authored list
scrolling, actions, shared/orphaned widgets, multiselection and ambiguous labels
or exports are skipped before mutation. This replaces original artwork; source
typography, borders and styling are not preserved or certified. Default choice
updates still request viewer regeneration. Text and choice appearance options
are independent; both can be selected in the demo's form panel.

The selected font must be `Helvetica`, `Times-Roman`, or `Courier`. This opts
into pdf-lib's default text appearance provider and replaces the original font
and appearance artwork. It supports printable ASCII (U+0020–U+007E), empty text,
and plain single-line text fields only. It is not an original-font fidelity or
universal rendering guarantee; fixed font sizes and long values may clip.
Rich text, multiline, comb, password, file selection, field/widget actions,
missing page widgets, and ambiguous/shared widget ownership are skipped before
the value changes. No unsupported character is replaced, transliterated, or dropped.

`applied` and `skipped` describe value updates. In opt-in mode, `appearances`
reports each applied text/choice request as `generated` (with the number of
widgets updated) or `viewer-required`. Checkbox/radio behavior
is unchanged. Every owned qualified widget is regenerated; unrelated fields retain
their appearances. Existing `NeedAppearances=true` is preserved because other
fields may still need regeneration. An unexpected generation failure rejects
the entire operation, returning no partially updated PDF. Source bytes are
never modified. Reopen and inspect both values and rendered widgets before
using an exported form.

The implementation uses pdf-lib's documented
[text field appearance API](https://pdf-lib.js.org/docs/api/classes/pdftextfield#updateappearances).

For an explicitly supplied embedded font, pass `textAppearance: { fontBytes }`
instead of `{ font: 'Helvetica' }`. `fontBytes` is a `Uint8Array` containing a
fixed TrueType outline face, standalone or selected from a TTC/OTC collection (up to 16 MiB). Pass zero-based `faceIndex` to select a collection face; it defaults to zero. The saved PDF contains a
complete font program and source-semantic Unicode mappings; reopening does not
require that font to be installed. The playground offers a local font picker and collection face index. Only the selected complete face is saved.

This single-line cluster profile accepts up to 4,096 UTF-16 units of Unicode text covered by the supplied font and HarfBuzz. Unicode 17 script extensions and grapheme segmentation preserve contextual script clusters. Unicode 17 bidi resolution handles mixed direction text, digits, brackets and isolates. Optional `direction` (`auto`, `ltr`, `rtl`) and `language` (default `und`) configure shaping. HarfBuzz
positions combining marks and ordinary ligatures, retaining exact source `/V`.
Distinct source strings sharing a glyph receive distinct CIDs. Unpartitionable
continuation glyphs use exact font outlines so reader extraction does not invent
characters; those outlines do not receive TrueType hinting. Generic readers may reorder RTL extraction (including lam-alef) despite exact `/V`, ToUnicode and ActualText. Auto-sized RTL and marks
fit positioned ink bounds; explicitly fixed sizes can still clip.

CFF font packaging and explicit font fallback remain subsequent milestones. Missing glyphs, invalid paths/metrics,
unsupported font formats and resource limits produce explicit refusal. A single
unpartitionable cluster may use at most 256 UTF-16 units (the PDF ToUnicode
512-byte destination limit). Existing ownership, source-preservation and mixed
standard-choice guards remain in force. This explicitly replaces the original
font. See [Unicode shaping and extraction policy](../../docs/PDF-UNICODE-SHAPING.md)
for the current scope, reader behavior and remaining work.
