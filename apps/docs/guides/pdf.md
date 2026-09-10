# PDFs

Use `@injoffice/pdf/browser` for byte-oriented page operations and browser viewing. The root entry includes additional Node-qualified integrations; importing it indiscriminately into a web application can pull in unavailable host APIs.

## Rotate and reopen

<<< @/examples/pdf.ts

This example preserves the input buffer, rotates the final page, and reparses the output. PDF selectors are **one-based**. `pages: [1]` targets the first page; do not reuse zero-based spreadsheet coordinates.

The returned `before` and `after` metadata let your host check the requested page rotation and relevant page geometry in addition to page count.

## Add a browser viewer

Your bundler must emit a matching PDF.js Worker. For Vite, the package reference shows:

```ts
import workerURL from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { configurePdfWorker, PdfViewerDocument } from '@injoffice/pdf/browser'

configurePdfWorker(workerURL)
// source is the Uint8Array supplied by your host's file-selection flow.
const viewer = await PdfViewerDocument.load(source)
```

This is a Vite-specific integration fragment, not a standalone Node script. Install `pdfjs-dist` explicitly if your application imports it, and align its version with the PDF package's supported dependency. Follow viewer lifecycle methods in your installed declarations when unmounting.

## Agent page operations

The PDF agent adapter advertises rotate, blank-page insertion, delete, reorder, crop, resize, and n-up operations. It does not expose arbitrary PDF text rewriting or redaction. [Try a complete reviewable agent workflow](../agents/quickstart).

## Destructive and Node-only tools

The root package contains additional annotation, form, image, OCR, and redaction paths. Consult the [package reference](../reference/generated/packages/pdf) and exact exports for runtime requirements.

Redaction requires stronger acceptance checks than hiding content visually or successfully reparsing the file. Do not equate a crop box, black overlay, or preview with removal of sensitive content. Apply the specific operation's verification contract and your host's security policy before releasing output.

Treat all input and transformed files as untrusted: enforce byte/page/resource limits and keep parsing failures separate from verified success.
