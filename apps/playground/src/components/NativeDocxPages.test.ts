import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { NativeDocxPages, NativeDocxImage, nativeDocxImageOrientation, decodeNativeDocxImages, nativeDocxImagesWithinBudget, nativeDocxSVGPath, readNativePreviewResponse,nativeDocxFontSubstitutionSummary } from './NativeDocxPages'

describe('native document page viewer', () => {
  it('keeps the persistent font warning and deduplicated selections separate from technical evidence',()=>{
    const record={source_family:'Missing',selected_family:'Selected',weight:400,style:'normal'}
    const summary=nativeDocxFontSubstitutionSummary({substitutions:[record,record]} as any)
    expect(summary).toHaveLength(2)
    expect(summary[0]).toContain('layout may differ')
    expect(summary[1]).toBe('Missing → Selected (400, normal).')
    expect(summary.join(' ')).not.toContain('sha256:')
  })
  it('offers an explicit separate font-substitution upload without enabling it automatically',()=>{
    const html=renderToStaticMarkup(createElement(NativeDocxPages,{bytes:new Uint8Array([1]),packageDigest:`sha256:${'a'.repeat(64)}`,apiBase:'https://helper.invalid'}))
    expect(html).toContain('Upload to helper and allow operator font substitution')
    expect(html).toContain('cannot combine with the other approximate page policies')
    expect(html).not.toContain('Approximate page limitations')
  })
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
      expect(markup).toContain('Upload to helper and try approximate pages')
      expect(markup).toContain('does not reproduce older Word pagination')
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
  it('replays all source flips and half-turns as exact box-preserving matrices', () => {
    for (const rotation_degrees of [0, 180] as const) for (const flip_horizontal of [false, true]) for (const flip_vertical of [false, true]) {
      const sx = flip_horizontal !== (rotation_degrees === 180) ? -1 : 1
      const sy = flip_vertical !== (rotation_degrees === 180) ? -1 : 1
      const markup = renderToStaticMarkup(createElement(NativeDocxImage, { command: { kind: 'paint_inline_image', id: 'image:1', line_id: 'line:1', fragment_id: 'fragment:1', source_id: 'run:1', drawing_id: 'drawing:1', asset_id: 'asset:1', x_millipoints: 10, y_millipoints: 20, width_millipoints: 200, height_millipoints: 100, source_crop: { left: 0, top: 0, right: 0, bottom: 0, unit: 'one-hundred-thousandth' }, transform: { rotation_degrees, flip_horizontal, flip_vertical } }, base64: 'AA==' }))
      expect(markup).toContain(`transform="matrix(${sx} 0 0 ${sy} ${sx < 0 ? 220 : 0} ${sy < 0 ? 140 : 0})"`)
      expect(markup).toContain('width="200" height="100"')
    }
  })
  it('maps every oriented source corner to the exact attested box for all quarter turns and flips', () => {
    for (const rotation_degrees of [0, 90, 180, 270] as const) for (const flip_horizontal of [false, true]) for (const flip_vertical of [false, true]) {
      const geometry = nativeDocxImageOrientation({ x_millipoints: 10, y_millipoints: 20, width_millipoints: 200, height_millipoints: 100, transform: { rotation_degrees, flip_horizontal, flip_vertical } } as Parameters<typeof nativeDocxImageOrientation>[0])
      const [a,b,c,d,e,f] = geometry.matrix as [number, number, number, number, number, number]
      for (const u of [0,1]) for (const v of [0,1]) {
        const sourceX = 10 + u * geometry.width, sourceY = 20 + v * geometry.height
        const reflectedU = flip_horizontal ? 1-u : u, reflectedV = flip_vertical ? 1-v : v
        const expected = rotation_degrees === 90 ? [1-reflectedV,reflectedU] : rotation_degrees === 180 ? [1-reflectedU,1-reflectedV] : rotation_degrees === 270 ? [reflectedV,1-reflectedU] : [reflectedU,reflectedV]
        expect([a*sourceX+c*sourceY+e,b*sourceX+d*sourceY+f]).toEqual([10+expected[0]!*200,20+expected[1]!*100])
      }
    }
  })
  it('clips the original source rectangle before applying image orientation', () => {
    const markup = renderToStaticMarkup(createElement(NativeDocxImage, { command: { kind: 'paint_inline_image', id: 'image:1', line_id: 'line:1', fragment_id: 'fragment:1', source_id: 'run:1', drawing_id: 'drawing:1', asset_id: 'asset:1', x_millipoints: 10, y_millipoints: 20, width_millipoints: 200, height_millipoints: 100, source_crop: { left: 50000, top: 0, right: 0, bottom: 0, unit: 'one-hundred-thousandth' }, transform: { rotation_degrees: 90, flip_horizontal: false, flip_vertical: false } }, base64: 'AA==' }))
    expect(markup).toContain('viewBox="50000 0 50000 100000"')
    expect(markup).toContain('overflow="hidden"')
    expect(markup).toContain('width="100" height="200"')
    expect(markup).toContain('<image x="0" y="0" width="100000" height="100000"')
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
