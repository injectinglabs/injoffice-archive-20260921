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
