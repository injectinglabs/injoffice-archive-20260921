import { describe, expect, it, vi } from 'vitest';
import { PDFDocument, PDFName } from 'pdf-lib';
import { paintLinkBorders } from './linkBorders.js';
import { PdfViewerDocument, renderPageToCanvas } from './viewer.js';

function context() {
  return { save: vi.fn(), restore: vi.fn(), setTransform: vi.fn(), transform: vi.fn(), setLineDash: vi.fn(), beginPath: vi.fn(), rect: vi.fn(), ellipse: vi.fn(), closePath: vi.fn(), moveTo: vi.fn(), lineTo: vi.fn(), stroke: vi.fn() } as unknown as CanvasRenderingContext2D;
}
const link = { subtype: 'Link', annotationFlags: 0, hasAppearance: false, rect: [10, 20, 110, 60], borderColor: new Uint8ClampedArray([255, 0, 0]), borderStyle: { width: 2, style: 1, dashArray: [3, 2] } };
const transform = [0, 2, 2, 0, 40, -20];

describe('canvas link annotation borders', () => {
  it('paints solid rectangles with the page rotation and backing-store scale', () => {
    const ctx = context();
    paintLinkBorders(ctx, [link], transform, 1.5);
    expect(ctx.setTransform).toHaveBeenCalledWith(1.5, 0, 0, 1.5, 0, 0);
    expect(ctx.transform).toHaveBeenCalledWith(...transform);
    expect(ctx.rect).toHaveBeenCalledWith(10, 20, 100, 40);
    expect(ctx.strokeStyle).toBe('rgb(255, 0, 0)');
    expect(ctx.stroke).toHaveBeenCalledOnce();
    expect(ctx.restore).toHaveBeenCalledOnce();
  });

  it('preserves dash lengths and paints underline at the PDF bottom edge', () => {
    const ctx = context();
    paintLinkBorders(ctx, [
      { ...link, borderStyle: { ...link.borderStyle, style: 2 } },
      { ...link, borderStyle: { ...link.borderStyle, style: 5 } },
    ], transform, 1);
    expect(ctx.setLineDash).toHaveBeenNthCalledWith(1, [3, 2]);
    expect(ctx.setLineDash).toHaveBeenNthCalledWith(2, []);
    expect(ctx.moveTo).toHaveBeenCalledWith(10, 20);
    expect(ctx.lineTo).toHaveBeenCalledWith(110, 20);
  });

  it('uses portable ellipse paths for rounded borders', () => {
    const ctx = context();
    paintLinkBorders(ctx, [{ ...link, borderStyle: { ...link.borderStyle, horizontalCornerRadius: 5, verticalCornerRadius: 3 } }], transform, 1);
    expect(ctx.ellipse).toHaveBeenCalledTimes(4);
    expect(ctx.closePath).toHaveBeenCalledOnce();
    expect(ctx.stroke).toHaveBeenCalledOnce();
  });

  it('does not double-paint appearances, hidden annotations, or unsupported geometry', () => {
    const ctx = context();
    const variations = [{ hasAppearance: true }, { annotationFlags: 1 }, { annotationFlags: 2 }, { annotationFlags: 32 }, { subtype: 'Widget' }, { quadPoints: [0] }, { noRotate: true }, { noZoom: true }, { borderColor: null }, { rect: [0, 0, Infinity, 10] }, { borderStyle: { width: 0, style: 1 } }];
    paintLinkBorders(ctx, variations.map((v) => ({ ...link, ...v })), transform, 1);
    expect(ctx.stroke).not.toHaveBeenCalled();
  });

  it('uses PDF.js annotation data from a real generated PDF', async () => {
    const pdf = await PDFDocument.create();
    const page = pdf.addPage([200, 200]);
    page.node.set(PDFName.of('Annots'), pdf.context.obj([pdf.context.register(pdf.context.obj({
      Type: 'Annot', Subtype: 'Link', Rect: [10, 20, 110, 60], C: [1, 0, 0], BS: { W: 2, S: 'D', D: [3, 2] },
    }))]));
    const doc = await PdfViewerDocument.load(await pdf.save());
    try {
      const annotations = await (await doc.getPage(1)).getAnnotations({ intent: 'display' });
      const ctx = context();
      paintLinkBorders(ctx, annotations, transform, 1);
      expect(ctx.stroke).toHaveBeenCalledOnce();
      expect(ctx.setLineDash).toHaveBeenCalledWith([3, 2]);
    } finally { await doc.destroy(); }
  });

  it('adds borders after the page canvas finishes rendering', async () => {
    const ctx = context();
    const getAnnotations = vi.fn(async () => [link]);
    const page = { getViewport: () => ({ width: 100, height: 200, transform }), render: () => ({ promise: Promise.resolve(), cancel: vi.fn() }), getAnnotations };
    const doc = { getPage: async () => page } as unknown as PdfViewerDocument;
    const canvas = { getContext: () => ctx, style: {} } as unknown as HTMLCanvasElement;
    await renderPageToCanvas(doc, 1, canvas);
    expect(getAnnotations).toHaveBeenCalledWith({ intent: 'display' });
    expect(ctx.stroke).toHaveBeenCalledOnce();
  });
});
