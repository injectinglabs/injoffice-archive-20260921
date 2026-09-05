/** Straight-alpha BGRA (pdfium's raster pixel format) → PNG bytes. Mirrors decodePng.ts's
 * RGBA→BGRA swap in reverse; no premultiply handling either way since PNG straight alpha
 * is what pdfium's opaque page raster already is. */
import { PNG } from 'pngjs';

export function encodePng(width: number, height: number, bgra: Buffer): Uint8Array {
  const png = new PNG({ width, height });
  for (let i = 0; i < bgra.length; i += 4) {
    png.data[i] = bgra[i + 2]!;
    png.data[i + 1] = bgra[i + 1]!;
    png.data[i + 2] = bgra[i]!;
    png.data[i + 3] = bgra[i + 3]!;
  }
  return PNG.sync.write(png);
}
