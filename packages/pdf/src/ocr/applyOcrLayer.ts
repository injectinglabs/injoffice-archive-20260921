/**
 * Splice a synthetic, invisible text layer onto pages that have none — the OCR pipeline's
 * write side. Strictly additive: a page with any real extracted text is never touched, only
 * rendered and handed to `provider` when it has none. Each recognized word becomes its own
 * invisible (`Tr 3`) text run, positioned and horizontally scaled (`Tz`) to its reported box
 * so search/copy/screen-reader access line up with what's on the page — the same idea real
 * OCR products use, built against pdf-lib's public operator API.
 */
import { PDFDocument, PDFFont, PDFPage, StandardFonts, popGraphicsState, pushGraphicsState, setCharacterSqueeze, setTextRenderingMode, TextRenderingMode } from 'pdf-lib';
import { PdfViewerDocument } from '../viewer.js';
import { renderPageToPng } from './renderPage.js';
import type { OcrLayerResult, OcrPageFailure, OcrProvider, OcrWord } from './types.js';

/** Places one word as invisible text sized/scaled to its box. Returns false (no-op, not a
 * failure) for a blank word, a degenerate box, or text this font's encoding can't represent
 * — OCR output is untrusted, and one bad word should reduce coverage, not abort the page. */
function spliceWord(page: PDFPage, font: PDFFont, word: OcrWord, pageWidthPt: number, pageHeightPt: number): boolean {
  const text = word.text.trim();
  if (!text) return false;
  const [bx1, by1, bx2, by2] = word.box.map((v) => Math.min(1, Math.max(0, v))) as [number, number, number, number];
  const boxW = (bx2 - bx1) * pageWidthPt;
  const boxH = (by2 - by1) * pageHeightPt;
  if (boxW <= 0 || boxH <= 0) return false;
  const x = bx1 * pageWidthPt;
  const y = (1 - by2) * pageHeightPt;
  const fontSize = Math.min(400, Math.max(1, boxH * 0.85));
  try {
    const naturalWidth = font.widthOfTextAtSize(text, fontSize);
    const scalePct = naturalWidth > 0 ? Math.min(500, Math.max(1, (boxW / naturalWidth) * 100)) : 100;
    page.pushOperators(pushGraphicsState(), setTextRenderingMode(TextRenderingMode.Invisible), setCharacterSqueeze(scalePct));
    page.drawText(text, { x, y, size: fontSize, font });
    page.pushOperators(popGraphicsState());
    return true;
  } catch {
    return false;
  }
}

/**
 * Render each page with no existing text layer, hand it to `provider`, and splice the
 * recognized words back in as invisible text. Fail-soft per page: a provider error (or a
 * page that yields zero placeable words) is reported in `failed`/omitted from `pagesOcred`
 * without blocking the rest of the document.
 */
export async function applyOcrLayer(bytes: Uint8Array, provider: OcrProvider, scale = 2): Promise<OcrLayerResult> {
  const viewer = await PdfViewerDocument.load(bytes);
  try {
    const pagesOcred: number[] = [];
    const pagesSkippedExistingText: number[] = [];
    const failed: OcrPageFailure[] = [];
    const pdfDoc = await PDFDocument.load(bytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    let wroteAny = false;
    for (let page = 1; page <= viewer.pageCount; page++) {
      const existingText = await viewer.getPageText(page);
      if (existingText.trim().length > 0) {
        pagesSkippedExistingText.push(page);
        continue;
      }
      try {
        const { width: pageWidthPt, height: pageHeightPt } = await viewer.getPageSize(page);
        const rendered = await renderPageToPng(bytes, page, scale);
        const words = await provider({ page, png: rendered.png, width: rendered.width, height: rendered.height });
        const pdfPage = pdfDoc.getPages()[page - 1]!;
        let wrote = false;
        for (const word of words) {
          if (spliceWord(pdfPage, font, word, pageWidthPt, pageHeightPt)) wrote = true;
        }
        if (wrote) {
          pagesOcred.push(page);
          wroteAny = true;
        } else {
          failed.push({ page, reason: 'no placeable words returned' });
        }
      } catch (err) {
        failed.push({ page, reason: err instanceof Error ? err.message : String(err) });
      }
    }
    return { bytes: wroteAny ? await pdfDoc.save() : bytes, pagesOcred, pagesSkippedExistingText, failed };
  } finally {
    await viewer.destroy();
  }
}
