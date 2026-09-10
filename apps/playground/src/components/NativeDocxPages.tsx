import { useEffect, useRef, useState } from 'react'
import type { NativeDocxPagePaintV1, NativeDocxPaintPathCommandV1, NativeDocxPaintInlineImageCommandV1 } from '../../../../packages/docs/src/nativePagePaintV1'
import { decodeNativeDocxPagePaintV1 } from '../../../../packages/docs/src/nativePagePaintOutputV1'
import { DsButton } from '../design-system/primitives'

export function NativeDocxImage({ command, base64 }: { command: NativeDocxPaintInlineImageCommandV1; base64: string }) {
  return <image x={command.x_millipoints} y={command.y_millipoints} width={command.width_millipoints} height={command.height_millipoints} preserveAspectRatio="none" href={`data:image/png;base64,${base64}`} />
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
  const [paint, setPaint] = useState<NativeDocxPagePaintV1 | null>(null)
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
  async function render() {
    const token = ++generation.current
    const controller = new AbortController(); pending.current = controller
    setBusy(true); setPaint(null); setMessage('Compiling native pages using document-bound fonts…')
    try {
      const response = await fetch(`${apiBase}/v1/docx/page-preview`, { method: 'POST', headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, body: new Blob([Uint8Array.from(bytes).buffer]), signal: controller.signal, credentials: 'omit', redirect: 'error' })
      const value = await readNativePreviewResponse(response) as Record<string, unknown>
      if (!response.ok) throw new Error(typeof value.error === 'string' ? value.error : 'Native preview was refused by the helper.')
      const decoded = decodeNativeDocxPagePaintV1(value.page_paint_output)
      if (!decoded.ok || value.canonical_output_validated !== true) throw new Error('Native preview failed schema validation.')
      if (decoded.value.provenance.package_sha256 !== packageDigest) throw new Error('Native pages do not match the currently opened document.')
      if (token !== generation.current) return
      if (!nativeDocxImagesWithinBudget(decoded.value.resources)) throw new Error('Native images exceed the interactive viewer pixel budget.')
      if (decoded.value.status === 'painted' && decoded.value.pages.some((page) => page.commands.length > 20_000 || page.commands.reduce((count, command) => count + (command.kind === 'fill_glyph_path' ? command.path.length : 0), 0) > 200_000)) throw new Error('Native page geometry exceeds the interactive viewer budget.')
      setPaint(decoded.value); setPageIndex(0)
      setMessage(decoded.value.status === 'painted' ? `${decoded.value.pages.length} native page${decoded.value.pages.length === 1 ? '' : 's'}. Read-only native page geometry. Supported text edits, when available, are offered in the content preview below.` : 'Native page rendering refused this document. The approximate content preview below remains available; the original file is unchanged.')
    } catch (error) {
      if (token !== generation.current || controller.signal.aborted) return
      setMessage(`${error instanceof Error ? error.message : 'Native preview failed.'} The approximate content preview remains available; the original file is unchanged.`)
    } finally { if (token === generation.current) setBusy(false) }
  }
  return <section className="docx-native-pages" aria-label="Native document pages">
    <h3>Native page preview</h3>
    <p role="status">{message}</p>
    <DsButton disabled={busy} onClick={() => void render()}>{busy ? 'Rendering native pages…' : 'Upload to helper and render native pages'}</DsButton>
    {paint?.status === 'painted' && <nav aria-label="Native document page navigation"><DsButton disabled={pageIndex === 0} onClick={() => setPageIndex((index) => index - 1)}>Previous native page</DsButton><span>Page {pageIndex + 1} of {paint.pages.length}</span><DsButton disabled={pageIndex >= paint.pages.length - 1} onClick={() => setPageIndex((index) => index + 1)}>Next native page</DsButton></nav>}
    {paint?.status === 'painted' && paint.pages.slice(pageIndex, pageIndex + 1).map((page) => <figure key={page.id}>
      <svg role="img" aria-label={`Native document page ${page.ordinal + 1}`} viewBox={`0 0 ${page.width_millipoints} ${page.height_millipoints}`} style={{ display: 'block', width: '100%', maxWidth: `${page.width_millipoints / 750}px`, background: '#fff', border: '1px solid var(--ds-border, #d5d9df)' }}>
        {page.commands.map((command) => {
          switch (command.kind) {
            case 'fill_glyph_path': return <path key={command.id} d={nativeDocxSVGPath(command.path)} fill={`#${command.fill_rgb}`} fillRule="nonzero" />
            case 'fill_table_cell': return <rect key={command.id} x={command.x_millipoints} y={command.y_millipoints} width={command.width_millipoints} height={command.height_millipoints} fill={`#${command.fill_rgb}`} />
            case 'stroke_table_border': case 'stroke_note_separator': return <line key={command.id} x1={command.x1_millipoints} y1={command.y1_millipoints} x2={command.x2_millipoints} y2={command.y2_millipoints} stroke={`#${command.stroke_rgb}`} strokeWidth={command.width_millipoints} />
            case 'paint_inline_image': { const asset = paint.resources.find((asset) => asset.id === command.asset_id); return asset ? <NativeDocxImage key={command.id} command={command} base64={asset.bytes_base64} /> : null }
          }
        })}
      </svg><figcaption>Page {page.ordinal + 1}</figcaption>
    </figure>)}
  </section>
}
