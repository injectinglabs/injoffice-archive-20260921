import { PDFDocument, rgb } from 'pdf-lib';
import { PNG } from 'pngjs';
import { describe, expect, it } from 'vitest';
import { renderPageToPng } from './renderPage.js';

async function docWithFilledRect(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([200, 100]);
  page.drawRectangle({ x: 20, y: 20, width: 60, height: 40, color: rgb(1, 0, 0) });
  return doc.save();
}

describe('renderPageToPng', () => {
  it('rasterizes a page at the requested scale and produces a decodable PNG', async () => {
    const bytes = await docWithFilledRect();
    const rendered = await renderPageToPng(bytes, 1, 2);
    expect(rendered.width).toBe(400);
    expect(rendered.height).toBe(200);
    const png = PNG.sync.read(Buffer.from(rendered.png));
    expect(png.width).toBe(400);
    expect(png.height).toBe(200);
  });

  it('paints the page background white and the drawn content in its real color', async () => {
    const bytes = await docWithFilledRect();
    const rendered = await renderPageToPng(bytes, 1, 1);
    const png = PNG.sync.read(Buffer.from(rendered.png));
    // top-left corner is outside the rect (page y-up, PNG y-down: rect is near the bottom)
    const corner = (png.width * 2 + 2) << 2;
    expect([png.data[corner], png.data[corner + 1], png.data[corner + 2]]).toEqual([255, 255, 255]);
    // a point inside the drawn rectangle: rect is x:[20,80] y:[20,60] in PDF space (y-up),
    // page height 100 -> in image space (y-down) that's rows [40,80]
    const inside = ((60 * png.width + 50) << 2);
    expect(png.data[inside]).toBeGreaterThan(200);
    expect(png.data[inside + 1]).toBeLessThan(60);
    expect(png.data[inside + 2]).toBeLessThan(60);
  });

  it('rejects an out-of-range page', async () => {
    const bytes = await docWithFilledRect();
    await expect(renderPageToPng(bytes, 5, 1)).rejects.toThrow(/does not exist/);
  });
});
