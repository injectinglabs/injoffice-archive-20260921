import { useEffect, useRef, useState } from 'react'
import type { NativeDocxPagePaintV1, NativeDocxPaintPathCommandV1, NativeDocxPaintInlineImageCommandV1, NativeDocxPaintFloatingImageCommandV1 } from '../../../../packages/docs/src/nativePagePaintV1'
import { decodeNativeDocxPagePaintV1 } from '../../../../packages/docs/src/nativePagePaintOutputV1'
import { decodeNativeDocxApproximatePagePreviewV1, decodeNativeDocxAutomaticBorderPreviewV1, DOCX_AUTO_BORDER_PREVIEW_PROTOCOL } from '@injoffice/docs/native-page-paint-output'
import { DsButton } from '../design-system/primitives'

export function NativeDocxImage({ command, base64, contentType = 'image/png', onError }: { command: NativeDocxPaintInlineImageCommandV1 | NativeDocxPaintFloatingImageCommandV1; base64: string; contentType?: 'image/png' | 'image/jpeg'; onError?: () => void }) {
  const geometry = nativeDocxImageOrientation(command)
  const crop = command.source_crop
  if (crop.left || crop.top || crop.right || crop.bottom) return <svg data-native-crop="true" x={command.x_millipoints} y={command.y_millipoints} width={geometry.width} height={geometry.height} viewBox={`${crop.left} ${crop.top} ${100000-crop.left-crop.right} ${100000-crop.top-crop.bottom}`} transform={`matrix(${geometry.matrix.join(' ')})`} preserveAspectRatio="none" overflow="hidden"><image x={0} y={0} width={100000} height={100000} preserveAspectRatio="none" href={`data:${contentType};base64,${base64}`} onError={onError} /></svg>
  return <image x={command.x_millipoints} y={command.y_millipoints} width={geometry.width} height={geometry.height} transform={`matrix(${geometry.matrix.join(' ')})`} preserveAspectRatio="none" href={`data:${contentType};base64,${base64}`} onError={onError} />
}

/** Reflect in source axes, then rotate clockwise into the attested inline box.
 * Integer matrices avoid transform-origin and trigonometric rounding. */
export function nativeDocxImageOrientation(command: NativeDocxPaintInlineImageCommandV1 | NativeDocxPaintFloatingImageCommandV1) {
  const { x_millipoints: x, y_millipoints: y, width_millipoints: boxWidth, height_millipoints: boxHeight } = command
  const angle = command.transform.rotation_degrees, quarter = angle === 90 || angle === 270
  const width = quarter ? boxHeight : boxWidth, height = quarter ? boxWidth : boxHeight
  const sx = command.transform.flip_horizontal ? -1 : 1, sy = command.transform.flip_vertical ? -1 : 1
  const dx = sx < 0 ? width : 0, dy = sy < 0 ? height : 0
  const local = angle === 90 ? [0, sx, -sy, 0, x + height - dy, y + dx]
    : angle === 180 ? [-sx, 0, 0, -sy, x + width - dx, y + height - dy]
      : angle === 270 ? [0, -sx, sy, 0, x + dy, y + width - dx]
        : [sx, 0, 0, sy, x + dx, y + dy]
  const [a, b, c, d, e, f] = local as [number, number, number, number, number, number]
  return { width, height, matrix: [a, b, c, d, e - a*x - c*y, f - b*x - d*y].map((value) => value === 0 ? 0 : value) }
}

/** Marker admission is not pixel decoding. Check every bounded asset before
 * claiming a painted page, one image at a time with a total time budget. */
export async function decodeNativeDocxImages(resources: NativeDocxPagePaintV1['resources'], signal: AbortSignal, budgetMs = 15000): Promise<void> {
  const deadline = performance.now() + budgetMs
  for (const resource of resources) {
    signal.throwIfAborted()
    if (performance.now() >= deadline) throw new Error('Native image decoding exceeded its time budget.')
    const image = new Image()
    await new Promise<void>((resolve, reject) => {
      let settled = false
      const finish = (error?: unknown) => {
        if (settled) return
        settled = true; clearTimeout(timer); signal.removeEventListener('abort', abort)
        image.src = ''
        if (error) reject(error); else resolve()
      }
      const abort = () => finish(signal.reason ?? new Error('Native image decoding cancelled.'))
      const timer = setTimeout(() => finish(new Error('Native image decoding exceeded its time budget.')), Math.max(1, deadline - performance.now()))
      signal.addEventListener('abort', abort, { once: true })
      try {
        image.src = `data:${resource.content_type};base64,${resource.bytes_base64}`
        void image.decode().then(() => {
          if (image.naturalWidth !== resource.width_px || image.naturalHeight !== resource.height_px) finish(new Error('Native image decoded dimensions do not match the validated source.'))
          else finish()
        }, () => finish(new Error('Native image could not be decoded.')))
      } catch { finish(new Error('Native image could not be decoded.')) }
    })
  }
}

export function nativeDocxImagesWithinBudget(resources: readonly { width_px: number; height_px: number }[]): boolean {
  let pixels = 0
  for (const resource of resources) {
    const size = resource.width_px * resource.height_px
    if (!Number.isSafeInteger(size) || size < 1 || size > 16_000_000) return false
    pixels += size
    if (pixels > 32_000_000) return false
  }
  return true
}

export async function readNativePreviewResponse(response: Response, maxBytes = 16 * 1024 * 1024): Promise<unknown> {
  if (!response.body) throw new Error('Native preview response was empty.')
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []; let length = 0
  try {
    while (true) {
      const chunk = await reader.read(); if (chunk.done) break
      length += chunk.value.byteLength
      if (length > maxBytes) throw new Error('Native preview exceeded the response budget.')
      chunks.push(chunk.value)
    }
  } finally { await reader.cancel(); reader.releaseLock() }
  const bytes = new Uint8Array(length); let offset = 0
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length }
  return JSON.parse(new TextDecoder().decode(bytes))
}

export function nativeDocxSVGPath(commands: readonly NativeDocxPaintPathCommandV1[]): string {
  return commands.map((p) => {
    switch (p.kind) {
      case 'move_to': return `M${p.x_millipoints} ${p.y_millipoints}`
      case 'line_to': return `L${p.x_millipoints} ${p.y_millipoints}`
      case 'quadratic_to': return `Q${p.control_x_millipoints} ${p.control_y_millipoints} ${p.x_millipoints} ${p.y_millipoints}`
      case 'cubic_to': return `C${p.control_1_x_millipoints} ${p.control_1_y_millipoints} ${p.control_2_x_millipoints} ${p.control_2_y_millipoints} ${p.x_millipoints} ${p.y_millipoints}`
      case 'close_path': return 'Z'
    }
  }).join(' ')
}

export function NativeDocxPages({ bytes, packageDigest, apiBase }: { bytes: Uint8Array; packageDigest: string; apiBase: string }) {
  const [paint, setPaint] = useState<(Pick<NativeDocxPagePaintV1, 'status' | 'pages' | 'resources'> & { approximate: boolean; reasons: readonly string[] }) | null>(null)
  const consent = `Native pages require uploading this document to ${apiBase}. Nothing is uploaded until you choose the button below.`
  const [message, setMessage] = useState(consent)
  const [pageIndex, setPageIndex] = useState(0)
  const [busy, setBusy] = useState(false)
  const pending = useRef<AbortController | null>(null)
  const generation = useRef(0)
  useEffect(() => {
    generation.current++
    pending.current?.abort()
    setPaint(null); setBusy(false); setPageIndex(0)
    setMessage(consent)
    return () => { generation.current++; pending.current?.abort() }
  }, [bytes, packageDigest, apiBase])
  async function render(approximate = false) {
    const token = ++generation.current
    const controller = new AbortController(); pending.current = controller
    setBusy(true); setPaint(null); setMessage('Compiling native pages using document-bound fonts…')
    try {
      const response = await fetch(`${apiBase}/v1/docx/${approximate ? 'page-preview-approximate' : 'page-preview'}`, { method: 'POST', headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, body: new Blob([Uint8Array.from(bytes).buffer]), signal: controller.signal, credentials: 'omit', redirect: 'error' })
      const value = await readNativePreviewResponse(response) as Record<string, unknown>
      if (!response.ok) throw new Error(typeof value.error === 'string' ? value.error : 'Native preview was refused by the helper.')
      let next: NonNullable<typeof paint>
      if (approximate) {
        const decoded = value.protocol === DOCX_AUTO_BORDER_PREVIEW_PROTOCOL ? decodeNativeDocxAutomaticBorderPreviewV1(value) : decodeNativeDocxApproximatePagePreviewV1(value)
        if (!decoded.ok) throw new Error('Approximate preview failed schema validation.')
        if (decoded.value.source.package_sha256 !== packageDigest) throw new Error('Approximate pages do not match the currently opened document.')
        next = { ...decoded.value, approximate: true }
      } else {
        const decoded = decodeNativeDocxPagePaintV1(value.page_paint_output)
        if (!decoded.ok || value.canonical_output_validated !== true) throw new Error('Native preview failed schema validation.')
        if (decoded.value.provenance.package_sha256 !== packageDigest) throw new Error('Native pages do not match the currently opened document.')
        next = { ...decoded.value, approximate: false, reasons: [] }
      }
      if (token !== generation.current) return
      if (!nativeDocxImagesWithinBudget(next.resources)) throw new Error('Native images exceed the interactive viewer pixel budget.')
      if (next.status === 'painted' && next.pages.some((page) => page.commands.length > 20_000 || page.commands.reduce((count, command) => count + (command.kind === 'fill_glyph_path' ? command.path.length : 0), 0) > 200_000)) throw new Error('Native page geometry exceeds the interactive viewer budget.')
      await decodeNativeDocxImages(next.resources, controller.signal)
      if (token !== generation.current) return
      setPaint(next); setPageIndex(0)
      setMessage(next.status === 'painted' ? (approximate ? `${next.pages.length} approximate, read-only pages. Pagination may differ from Word; the original file is unchanged.` : `${next.pages.length} native page${next.pages.length === 1 ? '' : 's'}. Read-only native page geometry. Supported text edits, when available, are offered in the content preview below.`) : 'Page rendering refused this document. The approximate content preview below remains available; the original file is unchanged.')
    } catch (error) {
      if (token !== generation.current || controller.signal.aborted) return
      setMessage(`${error instanceof Error ? error.message : 'Native preview failed.'} The approximate content preview remains available; the original file is unchanged.`)
    } finally { if (token === generation.current) setBusy(false) }
  }
  const displayedGeneration = generation.current
  const imageFailed = () => {
    if (displayedGeneration !== generation.current) return
    setPaint(null)
    setMessage('Native image could not be displayed. Native pages were cleared. The approximate content preview remains available; the original file is unchanged.')
  }
  return <section className="docx-native-pages" aria-label="Native document pages">
    <h3>Native page preview</h3>
    <p role="status">{message}</p>
    <DsButton disabled={busy} onClick={() => void render()}>{busy ? 'Rendering native pages…' : 'Upload to helper and render native pages'}</DsButton>
    <DsButton disabled={busy} onClick={() => void render(true)}>Upload to helper and try approximate pages</DsButton>
    <p>Approximate pages may use current layout rules for eligible older Word settings, an explicit 11 pt host default where the source has no font size, and black automatic table borders on a source-qualified white background. This does not reproduce older Word pagination. Each applied policy is disclosed below. Other unsupported features remain refused; the original file is unchanged.</p>
    {paint?.approximate && <aside aria-label="Approximate page limitations"><strong>Approximate · read-only · not Word-validated</strong><ul>{paint.reasons.slice(0, 20).map((reason, index) => <li key={index}>{reason}</li>)}</ul>{paint.reasons.length > 20 && <p>{paint.reasons.length - 20} additional limitations.</p>}</aside>}
    {paint?.status === 'painted' && <nav aria-label={`${paint.approximate ? 'Approximate' : 'Native'} document page navigation`}><DsButton disabled={pageIndex === 0} onClick={() => setPageIndex((index) => index - 1)}>Previous {paint.approximate ? 'approximate' : 'native'} page</DsButton><span>Page {pageIndex + 1} of {paint.pages.length}</span><DsButton disabled={pageIndex >= paint.pages.length - 1} onClick={() => setPageIndex((index) => index + 1)}>Next {paint.approximate ? 'approximate' : 'native'} page</DsButton></nav>}
    {paint?.status === 'painted' && paint.pages.slice(pageIndex, pageIndex + 1).map((page) => <figure key={page.id}>
      <svg role="img" aria-label={`${paint.approximate ? 'Approximate' : 'Native'} document page ${page.ordinal + 1}`} viewBox={`0 0 ${page.width_millipoints} ${page.height_millipoints}`} style={{ display: 'block', width: '100%', maxWidth: `${page.width_millipoints / 750}px`, background: '#fff', border: '1px solid var(--ds-border, #d5d9df)' }}>
        {page.commands.map((command) => {
          switch (command.kind) {
            case 'fill_glyph_path': return <path key={command.id} d={nativeDocxSVGPath(command.path)} fill={`#${command.fill_rgb}`} fillRule="nonzero" />
            case 'fill_text_highlight': return <rect key={command.id} data-native-highlight="true" x={command.x_millipoints} y={command.y_millipoints} width={command.width_millipoints} height={command.height_millipoints} fill={`#${command.fill_rgb}`} />
            case 'fill_table_cell': return <rect key={command.id} x={command.x_millipoints} y={command.y_millipoints} width={command.width_millipoints} height={command.height_millipoints} fill={`#${command.fill_rgb}`} />
            case 'stroke_table_border': case 'stroke_note_separator': case 'stroke_text_underline': return <line key={command.id} data-native-underline={command.kind === 'stroke_text_underline' ? 'true' : undefined} x1={command.x1_millipoints} y1={command.y1_millipoints} x2={command.x2_millipoints} y2={command.y2_millipoints} stroke={`#${command.stroke_rgb}`} strokeWidth={command.width_millipoints} />
            case 'paint_floating_image':
            case 'paint_inline_image': { const asset = paint.resources.find((asset) => asset.id === command.asset_id); return asset ? <NativeDocxImage key={command.id} command={command} base64={asset.bytes_base64} contentType={asset.content_type} onError={imageFailed} /> : null }
          }
        })}
      </svg><figcaption>{paint.approximate ? 'Approximate, read-only · ' : ''}Page {page.ordinal + 1}</figcaption>
    </figure>)}
  </section>
}
