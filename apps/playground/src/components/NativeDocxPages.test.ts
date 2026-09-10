import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { NativeDocxPages, NativeDocxImage, decodeNativeDocxImages, nativeDocxImagesWithinBudget, nativeDocxSVGPath, readNativePreviewResponse } from './NativeDocxPages'

describe('native document page viewer', () => {
  const resources = [{ content_type: 'image/jpeg', bytes_base64: 'AA==', width_px: 2, height_px: 1 }] as Parameters<typeof decodeNativeDocxImages>[0]
  it('requires successful decoding with exact source dimensions', async () => {
    const released: string[] = []
    vi.stubGlobal('Image', class { naturalWidth = 2; naturalHeight = 1; set src(value: string) { released.push(value) }; decode() { return Promise.resolve() } })
    try {
      await expect(decodeNativeDocxImages(resources, new AbortController().signal)).resolves.toBeUndefined()
      expect(released.at(-1)).toBe('')
      await expect(decodeNativeDocxImages([{ ...resources[0]!, width_px: 3 }], new AbortController().signal)).rejects.toThrow(/dimensions/)
    } finally { vi.unstubAllGlobals() }
  })
  it('refuses browser decoder failures rather than reporting painted pages', async () => {
    vi.stubGlobal('Image', class { src = ''; decode() { return Promise.reject(new Error('bad entropy')) } })
    try { await expect(decodeNativeDocxImages(resources, new AbortController().signal)).rejects.toThrow(/could not be decoded/) }
    finally { vi.unstubAllGlobals() }
  })
  it('bounds stalled decoding and cancels stale work', async () => {
    vi.stubGlobal('Image', class { src = ''; decode() { return new Promise<void>(() => {}) } })
    try {
      await expect(decodeNativeDocxImages(resources, new AbortController().signal, 5)).rejects.toThrow(/time budget/)
      const controller = new AbortController(), promise = decodeNativeDocxImages(resources, controller.signal)
      controller.abort()
      await expect(promise).rejects.toMatchObject({ name: 'AbortError' })
    } finally { vi.unstubAllGlobals() }
  })
  it('does not upload on render and names the actual destination', () => {
    const fetch = vi.fn(); vi.stubGlobal('fetch', fetch)
    try {
      const markup = renderToStaticMarkup(createElement(NativeDocxPages, { bytes: new Uint8Array([1]), packageDigest: `sha256:${'0'.repeat(64)}`, apiBase: 'https://example.test/helper' }))
      expect(markup).toContain('https://example.test/helper')
      expect(markup).toContain('Nothing is uploaded until')
      expect(markup).toContain('Upload to helper and render native pages')
      expect(fetch).not.toHaveBeenCalled()
    } finally { vi.unstubAllGlobals() }
  })
  it('replays numeric native path coordinates without HTML injection', () => {
    expect(nativeDocxSVGPath([{ kind: 'move_to', x_millipoints: 5, y_millipoints: 6 }, { kind: 'quadratic_to', control_x_millipoints: 1, control_y_millipoints: 2, x_millipoints: 3, y_millipoints: 4 }, { kind: 'close_path' }])).toBe('M5 6 Q1 2 3 4 Z')
  })
  it('uses the document image extents rather than intrinsic aspect ratio', () => {
    const markup = renderToStaticMarkup(createElement(NativeDocxImage, { command: { kind: 'paint_inline_image', id: 'image:1', line_id: 'line:1', fragment_id: 'fragment:1', source_id: 'run:1', asset_id: 'asset:1', x_millipoints: 10, y_millipoints: 20, width_millipoints: 200, height_millipoints: 100, source_crop: { left: 0, top: 0, right: 0, bottom: 0, unit: 'one-hundred-thousandth' }, transform: { rotation_degrees: 0, flip_horizontal: false, flip_vertical: false } }, base64: 'AA==' }))
    expect(markup).toContain('preserveAspectRatio="none"')
    expect(markup).toContain('width="200" height="100"')
  })
  it('bounds decoded PNG pixels independently of compressed response bytes', () => {
    expect(nativeDocxImagesWithinBudget([{ width_px: 10000, height_px: 10000 }])).toBe(false)
    expect(nativeDocxImagesWithinBudget(Array.from({ length: 3 }, () => ({ width_px: 4000, height_px: 4000 })))).toBe(false)
    expect(nativeDocxImagesWithinBudget([{ width_px: 1920, height_px: 1080 }])).toBe(true)
  })
  it('cancels streamed responses before consuming an unbounded body', async () => {
    const cancel = vi.fn()
    const body = new ReadableStream<Uint8Array>({ pull(controller) { controller.enqueue(new Uint8Array(8)) }, cancel })
    await expect(readNativePreviewResponse(new Response(body), 10)).rejects.toThrow(/budget/)
    expect(cancel).toHaveBeenCalledOnce()
    await expect(readNativePreviewResponse(new Response('{"ok":true}'))).resolves.toEqual({ ok: true })
  })
})
