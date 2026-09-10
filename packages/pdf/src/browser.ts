// Browser-safe entry: inspection/page transforms and the pdf.js viewer. The
// root entry also exposes PDFium/font-file operations that require Node.js.
export * from './types.js'
export {
  decodePdfCollabOperation,
  encodePdfCollabOperation,
  type PdfCollabDecode,
  type PdfCollabOperation,
} from './collab.js'
export {
  applyPageOps,
  mergeDocuments,
  normalizeRotation,
  readInfo,
  resolveSelector,
  splitDocument,
} from './pageOps.js'
export {
  canvasRenderMetrics,
  configurePdfWorker,
  PdfViewerDocument,
  renderPageToCanvas,
  type CanvasRenderMetrics,
  type PdfOutlineItem,
  type PdfRenderOptions,
  type SearchMatch,
} from './viewer.js'
