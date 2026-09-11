import { beforeEach, describe, expect, it, vi } from 'vitest';
const getDocument = vi.hoisted(() => vi.fn());
vi.mock('pdfjs-dist/legacy/build/pdf.mjs', () => ({ getDocument, GlobalWorkerOptions: {} }));
import { PdfViewerDocument, type PdfLoadOptions } from './viewer.js';

describe('host PDF resources', () => {
  beforeEach(() => getDocument.mockReset());

  it('forwards only explicit resources, uses packed CMaps, and preserves caller bytes', async () => {
    const destroy = vi.fn().mockResolvedValue(undefined);
    getDocument.mockReturnValue({ promise: Promise.resolve({ numPages: 1 }), destroy });
    const bytes = new Uint8Array([1, 2, 3]);
    const options = { cMapUrl: '/app/maps/', standardFontDataUrl: '/app/fonts/', wasmUrl: '/app/wasm/', useSystemFonts: false };
    const doc = await PdfViewerDocument.load(bytes, { ...options, url: 'https://untrusted.invalid/' } as PdfLoadOptions);
    const parameters = getDocument.mock.calls[0][0];
    expect(parameters).toEqual({ data: bytes, ...options, cMapPacked: true });
    expect(parameters.data).not.toBe(bytes);
    parameters.data[0] = 9;
    expect(bytes[0]).toBe(1);
    await doc.destroy();
    expect(destroy).toHaveBeenCalledOnce();
  });

  it('keeps upstream policy unchanged when options are omitted', async () => {
    getDocument.mockReturnValue({ promise: Promise.resolve({ numPages: 1 }), destroy: vi.fn() });
    await PdfViewerDocument.load(new Uint8Array([1]));
    expect(Object.keys(getDocument.mock.calls[0][0])).toEqual(['data']);
  });

  it('rejects invalid host configuration before starting a worker', async () => {
    for (const key of ['cMapUrl', 'standardFontDataUrl', 'wasmUrl']) {
      for (const value of ['', ' ', '/no-trailing-slash', ' /space/', 42]) {
        await expect(PdfViewerDocument.load(new Uint8Array(), { [key]: value } as PdfLoadOptions)).rejects.toThrow(key);
      }
    }
    await expect(PdfViewerDocument.load(new Uint8Array(), { useSystemFonts: 'false' } as unknown as PdfLoadOptions)).rejects.toThrow('boolean');
    expect(getDocument).not.toHaveBeenCalled();
  });

  it('cleans up a failed load without hiding its original error', async () => {
    const error = new Error('invalid PDF');
    const destroy = vi.fn().mockRejectedValue(new Error('cleanup failed'));
    getDocument.mockReturnValue({ promise: Promise.reject(error), destroy });
    await expect(PdfViewerDocument.load(new Uint8Array())).rejects.toBe(error);
    expect(destroy).toHaveBeenCalledOnce();
  });
});
