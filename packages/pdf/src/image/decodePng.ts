/**
 * PNG → straight-alpha BGRA for FPDFBitmap_CreateEx (FPDFBitmap_BGRA).
 *
 * pngjs returns the file's native non-premultiplied RGBA. PDFium bitmaps are
 * BGRA, so only the R/B channels are swapped; alpha is left as stored.
 */
import { PNG } from 'pngjs';

export interface DecodedImage {
  width: number;
  height: number;
  bgra: Buffer;
}

export function decodePng(base64: string): DecodedImage {
  let png: PNG;
  try {
    png = PNG.sync.read(Buffer.from(base64, 'base64'));
  } catch (err) {
    throw new Error(`could not decode the image data: ${err instanceof Error ? err.message : String(err)}`);
  }
  const bgra = Buffer.from(png.data);
  for (let i = 0; i < bgra.length; i += 4) {
    const red = bgra[i]!;
    bgra[i] = bgra[i + 2]!;
    bgra[i + 2] = red;
  }
  return { width: png.width, height: png.height, bgra };
}
