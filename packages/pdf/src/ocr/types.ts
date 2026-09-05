/**
 * Vision-LLM OCR — P3. A from-scratch pipeline that feeds rendered page images
 * to a vision-LLM instead of a local OCR engine (docs/ROADMAP-PDF.md's P3
 * section). The vision call itself is the caller's responsibility — same
 * "pure core, thin shell" split as this repo's `ConnectorSpec`s: this package
 * owns page rasterization, box-to-content-stream placement, and the
 * additive-only policy; the actual network call to a model plugs in as
 * `OcrProvider`.
 */

/** One recognized word/run, with its box normalized to the rendered image (0-1,
 * top-left origin, y down — the natural coordinate space a vision model reports in,
 * resolution-independent of whatever raster scale was actually used). */
export interface OcrWord {
  text: string;
  /** [x1, y1, x2, y2], each 0-1 relative to the rendered page image. */
  box: [number, number, number, number];
}

export interface OcrPageInput {
  /** 1-indexed, matching this package's page-op convention. */
  page: number;
  png: Uint8Array;
  /** Raster pixel dimensions of `png` (informational — `OcrWord.box` is normalized). */
  width: number;
  height: number;
}

export type OcrProvider = (input: OcrPageInput) => Promise<OcrWord[]>;

export interface OcrPageFailure {
  page: number;
  reason: string;
}

export interface OcrLayerResult {
  bytes: Uint8Array;
  /** Pages that got a synthetic text layer written. */
  pagesOcred: number[];
  /** Pages the provider was never called for because they already have real text. */
  pagesSkippedExistingText: number[];
  /** Pages where the provider (or the splice) failed; no partial layer is written for these. */
  failed: OcrPageFailure[];
}
