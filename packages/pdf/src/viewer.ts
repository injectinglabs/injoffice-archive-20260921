import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist/legacy/build/pdf.mjs';

export interface SearchMatch {
  pageIndex: number;
  /** Character offset of the match within that page's concatenated text. */
  offset: number;
  snippet: string;
}

export interface CanvasRenderMetrics {
  /** Backing-store dimensions used for raster fidelity. */
  pixelWidth: number;
  pixelHeight: number;
  /** CSS dimensions used to keep the requested document zoom. */
  cssWidth: number;
  cssHeight: number;
  /** pdf.js canvas transform for the backing-store pixel ratio. */
  transform?: [number, number, number, number, number, number];
}

/**
 * Resolves CSS and backing-store sizes for a pdf.js viewport. Exported so
 * browser hosts can qualify high-DPI behavior without requiring a DOM canvas.
 */
export function canvasRenderMetrics(width: number, height: number, pixelRatio = 1): CanvasRenderMetrics {
  if (!Number.isFinite(width) || width <= 0 || !Number.isFinite(height) || height <= 0) {
    throw new Error('PDF viewport dimensions must be positive finite numbers');
  }
  if (!Number.isFinite(pixelRatio) || pixelRatio <= 0) {
    throw new Error('PDF canvas pixel ratio must be a positive finite number');
  }
  return {
    pixelWidth: Math.ceil(width * pixelRatio),
    pixelHeight: Math.ceil(height * pixelRatio),
    cssWidth: width,
    cssHeight: height,
    transform: pixelRatio === 1 ? undefined : [pixelRatio, 0, 0, pixelRatio, 0, 0],
  };
}

/**
 * Points pdf.js at a worker asset supplied by the browser host/bundler.
 *
 * Browser applications should call this once before loading a document. For
 * example, Vite hosts can import `pdf.worker.min.mjs?url` and pass the emitted
 * URL here. Keeping asset resolution in the host makes the worker work under a
 * non-root deployment base (such as GitHub Pages) without baking a CDN into
 * the library.
 */
export function configurePdfWorker(workerSrc: string | URL): void {
  const resolved = workerSrc.toString().trim();
  if (resolved.length === 0) throw new Error('pdf.js worker URL must not be empty');
  pdfjsLib.GlobalWorkerOptions.workerSrc = resolved;
}

/**
 * A loaded PDF ready for viewing/search — the read-only counterpart to
 * pageOps' mutating pdf-lib pipeline. Wraps pdfjs-dist so we can extract
 * layout/text without owning a canvas; actual page rasterization happens
 * host-side (DOM canvas in the browser), see `renderPageToCanvas`.
 */
export class PdfViewerDocument {
  private constructor(
    private readonly proxy: PDFDocumentProxy,
    private readonly task: ReturnType<typeof pdfjsLib.getDocument>,
  ) {}

  static async load(bytes: Uint8Array): Promise<PdfViewerDocument> {
    // pdfjs-dist transfers/detaches a `data` buffer it's handed — copy first so
    // callers can safely keep using their own `bytes` afterward (e.g. locate a
    // rect via the viewer, then pass the same bytes to applyTextEdits).
    const task = pdfjsLib.getDocument({ data: bytes.slice() });
    const proxy = await task.promise;
    return new PdfViewerDocument(proxy, task);
  }

  get pageCount(): number {
    return this.proxy.numPages;
  }

  async getPage(pageIndex: number): Promise<PDFPageProxy> {
    if (pageIndex < 1 || pageIndex > this.pageCount) {
      throw new Error(`page ${pageIndex} is out of range for a ${this.pageCount}-page document`);
    }
    return this.proxy.getPage(pageIndex);
  }

  async getPageSize(pageIndex: number): Promise<{ width: number; height: number }> {
    const page = await this.getPage(pageIndex);
    const viewport = page.getViewport({ scale: 1 });
    return { width: viewport.width, height: viewport.height };
  }

  async getPageText(pageIndex: number): Promise<string> {
    const page = await this.getPage(pageIndex);
    const content = await page.getTextContent();
    return content.items.map((item) => ('str' in item ? item.str : '')).join(' ');
  }

  async getOutline(): Promise<Array<{ title: string; pageIndex: number | null }>> {
    const outline = await this.proxy.getOutline();
    if (!outline) return [];
    const flatten = async (
      items: Awaited<ReturnType<PDFDocumentProxy['getOutline']>>,
    ): Promise<Array<{ title: string; pageIndex: number | null }>> => {
      const out: Array<{ title: string; pageIndex: number | null }> = [];
      for (const item of items ?? []) {
        let pageIndex: number | null = null;
        if (item.dest) {
          try {
            const dest = typeof item.dest === 'string' ? await this.proxy.getDestination(item.dest) : item.dest;
            const ref = dest?.[0];
            if (ref) pageIndex = (await this.proxy.getPageIndex(ref)) + 1;
          } catch {
            pageIndex = null;
          }
        }
        out.push({ title: item.title, pageIndex });
        out.push(...(await flatten(item.items)));
      }
      return out;
    };
    return flatten(outline);
  }

  /** Case-insensitive substring search across every page's extracted text. */
  async search(query: string): Promise<SearchMatch[]> {
    if (query.trim().length === 0) return [];
    const needle = query.toLowerCase();
    const matches: SearchMatch[] = [];
    for (let pageIndex = 1; pageIndex <= this.pageCount; pageIndex++) {
      const text = await this.getPageText(pageIndex);
      const haystack = text.toLowerCase();
      let from = 0;
      while (true) {
        const at = haystack.indexOf(needle, from);
        if (at === -1) break;
        matches.push({ pageIndex, offset: at, snippet: text.slice(Math.max(0, at - 20), at + needle.length + 20) });
        from = at + needle.length;
      }
    }
    return matches;
  }

  async destroy(): Promise<void> {
    await this.task.destroy();
  }
}

/**
 * Renders one page into a caller-supplied canvas. Browser-only (needs a real
 * CanvasRenderingContext2D) — Node test coverage stops at `PdfViewerDocument`'s
 * text/metadata surface above; this is exercised in-browser once wired into
 * a host app.
 */
export async function renderPageToCanvas(
  doc: PdfViewerDocument,
  pageIndex: number,
  canvas: HTMLCanvasElement,
  scale = 1,
  pixelRatio = typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1,
): Promise<void> {
  const page = await doc.getPage(pageIndex);
  const viewport = page.getViewport({ scale });
  const context = canvas.getContext('2d');
  if (!context) throw new Error('canvas 2d context unavailable');
  const metrics = canvasRenderMetrics(viewport.width, viewport.height, pixelRatio);
  canvas.width = metrics.pixelWidth;
  canvas.height = metrics.pixelHeight;
  canvas.style.width = `${metrics.cssWidth}px`;
  canvas.style.height = `${metrics.cssHeight}px`;
  await page.render({ canvas, canvasContext: context, viewport, transform: metrics.transform }).promise;
}
