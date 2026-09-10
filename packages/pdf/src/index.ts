export * from './types.js';
export { decodePdfCollabOperation, encodePdfCollabOperation, type PdfCollabDecode, type PdfCollabOperation } from './collab.js';
export { applyPageOps, mergeDocuments, normalizeRotation, readInfo, resolveSelector, splitDocument } from './pageOps.js';
export {
  canvasRenderMetrics,
  configurePdfWorker,
  PdfViewerDocument,
  renderPageToCanvas,
  type CanvasRenderMetrics,
  type PdfOutlineItem,
  type PdfRenderOptions,
  type SearchMatch,
} from './viewer.js';
export { applyTextEdits, editBlock, editText, type BlockEditSpec, type SurgicalEditSpec, type TextEditFailure, type TextEditSpec, type TextEditsResult } from './textEdit/index.js';
export { applyTextInserts, type TextInsertFailure, type TextInsertSpec, type TextInsertsResult } from './textEdit/index.js';
export {
  applyMarkups,
  applyDrawings,
  applyNoteEdits,
  applyFormValues,
  applyStamps,
  applySignatureStamps,
  applyAnnotDeletes,
  listRedactableFormFields,
  redactFormFields,
  VISUAL_SIGNATURE_CONTENT_PREFIX,
  type AnnotDeleteSpec,
  type DrawingSpec,
  type FormValueFailure,
  type FormValueSpec,
  type FormFieldRedactionSpec,
  type FormFieldsRedactionResult,
  type FormValuesResult,
  type MarkupSpec,
  type MarkupType,
  type NoteEditSpec,
  type NoteReplyTarget,
  type RedactableFormFieldRef,
  type SignatureStampSpec,
  type StampSpec,
} from './annotate/index.js';
export {
  applyImageEdits,
  listPageImages,
  verifyImageEdits,
  type ImageEditFailure,
  type ImageEditSpec,
  type ImageEditsResult,
  type ImageLayer,
  type PageImageRef,
} from './image/index.js';
export {
  applyOcrLayer,
  renderPageToPng,
  type OcrLayerResult,
  type OcrPageFailure,
  type OcrPageInput,
  type OcrProvider,
  type OcrWord,
  type RenderedPage,
} from './ocr/index.js';
export { applyWholeTextRedactions, type WholeTextRedaction, type RedactionProof, type RedactionResult, type RedactionRect } from './redaction/apply.js';
export { applyWholeImageRedactions, type WholeImageRedaction, type ImageRedactionProof } from './redaction/image.js';
export { applyWholeAnnotationRedactions, listWholeAnnotationRedactionTargets, type WholeAnnotationRedactionTarget } from './redaction/annotation.js';
export {
  applyWholePathRedactions,
  listPagePaths,
  type PagePathRef,
  type PathRedactionProof,
  type PathRedactionResult,
  type PdfRect,
  type WholePathRedaction,
} from './redaction/path.js';
export {
  applyWholeXObjectRedactions,
  listPageXObjects,
  type PageXObjectRef,
  type WholeXObjectRedaction,
  type XObjectRedactionProof,
  type XObjectRedactionResult,
} from './redaction/xobject.js';
