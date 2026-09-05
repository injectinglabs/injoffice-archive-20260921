/**
 * Headless page rasterization for the OCR pipeline (P3). A full-page raster is
 * needed to hand a vision-LLM something to read; this uses pdfium raw exports
 * (`FPDF_RenderPageBitmap` and friends) and this package's pdfium-wrapper
 * conventions (see textEdit/pdfium.ts).
 */
import { chainPdfium, loadPdfium, withDocument, type Pdfium } from '../textEdit/pdfium.js';
import { encodePng } from './encodePng.js';

const FPDF_BITMAP_BGRA = 4;
const FPDF_ANNOT = 1;

export interface RenderedPage {
  width: number;
  height: number;
  png: Uint8Array;
}

/**
 * Rasterize one page to PNG at `scale` × its PDF point size (1pt ≈ 1px at scale 1) —
 * a plain white-filled raster with annotations painted.
 */
export function renderPageToPng(bytes: Uint8Array, page: number, scale = 2): Promise<RenderedPage> {
  return chainPdfium(async () => {
    const m = await loadPdfium();
    return withDocument(m, bytes, async (doc) => {
      const pageIndex = page - 1;
      const pageCount = m._FPDF_GetPageCount(doc);
      if (pageIndex < 0 || pageIndex >= pageCount) {
        throw new Error(`page ${page} does not exist in a ${pageCount}-page document`);
      }
      const pdfPage = m._FPDF_LoadPage(doc, pageIndex);
      if (!pdfPage) throw new Error(`could not load page ${page}`);
      try {
        return renderLoadedPage(m, pdfPage, scale);
      } finally {
        m._FPDF_ClosePage(pdfPage);
      }
    });
  });
}

function renderLoadedPage(m: Pdfium, pdfPage: number, scale: number): RenderedPage {
  const w = Math.max(1, Math.round(m._FPDF_GetPageWidthF(pdfPage) * scale));
  const h = Math.max(1, Math.round(m._FPDF_GetPageHeightF(pdfPage) * scale));
  const bufPtr = m._malloc(w * h * 4);
  const bmp = m._FPDFBitmap_CreateEx(w, h, FPDF_BITMAP_BGRA, bufPtr, w * 4);
  if (!bmp) {
    m._free(bufPtr);
    throw new Error('FPDFBitmap_CreateEx failed');
  }
  try {
    m._FPDFBitmap_FillRect(bmp, 0, 0, w, h, 0xffffffff);
    m._FPDF_RenderPageBitmap(bmp, pdfPage, 0, 0, w, h, 0, FPDF_ANNOT);
    const stride = m._FPDFBitmap_GetStride(bmp);
    const buf = m._FPDFBitmap_GetBuffer(bmp);
    const bgra = Buffer.alloc(w * h * 4);
    for (let row = 0; row < h; row++) {
      bgra.set(m.HEAPU8.subarray(buf + row * stride, buf + row * stride + w * 4), row * w * 4);
    }
    return { width: w, height: h, png: encodePng(w, h, bgra) };
  } finally {
    m._FPDFBitmap_Destroy(bmp);
    m._free(bufPtr);
  }
}
