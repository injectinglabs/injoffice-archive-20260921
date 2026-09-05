import { PNG } from 'pngjs';
import { describe, expect, it } from 'vitest';
import { decodePng } from './decodePng.js';

function makePng(w: number, h: number, pixels: [number, number, number, number][]): string {
  const png = new PNG({ width: w, height: h });
  pixels.forEach(([r, g, b, a], i) => {
    const idx = i << 2;
    png.data[idx] = r;
    png.data[idx + 1] = g;
    png.data[idx + 2] = b;
    png.data[idx + 3] = a;
  });
  return PNG.sync.write(png).toString('base64');
}

describe('decodePng', () => {
  it('swaps R and B, leaving G and A untouched — RGBA in, BGRA out', () => {
    const b64 = makePng(1, 1, [[10, 20, 30, 200]]);
    const { width, height, bgra } = decodePng(b64);
    expect(width).toBe(1);
    expect(height).toBe(1);
    expect([...bgra]).toEqual([30, 20, 10, 200]);
  });

  it('preserves straight (non-premultiplied) alpha — no premultiply-correction applied', () => {
    // A translucent pixel whose stored color would look different if it had been
    // premultiplied and needed un-premultiplying; decodePng must hand it back exactly
    // as PNG stores it (straight alpha).
    const b64 = makePng(1, 1, [[200, 100, 50, 128]]);
    const { bgra } = decodePng(b64);
    expect([...bgra]).toEqual([50, 100, 200, 128]);
  });

  it('decodes a multi-pixel image row-major', () => {
    const b64 = makePng(2, 1, [
      [255, 0, 0, 255],
      [0, 255, 0, 255],
    ]);
    const { bgra } = decodePng(b64);
    expect([...bgra]).toEqual([0, 0, 255, 255, 0, 255, 0, 255]);
  });

  it('throws a descriptive error for undecodable data', () => {
    expect(() => decodePng('not-a-real-png')).toThrow(/could not decode/);
  });
});
