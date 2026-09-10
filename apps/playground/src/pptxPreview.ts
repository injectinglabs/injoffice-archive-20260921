import type { NativeAsset, NativeElement, NativePptxDeck } from '@injoffice/pptx-native'

/** Preview policy only. Never use these approximations to authorize a mutation. */
export function previewColor(value: string | undefined, fallback = '#ffffff'): string {
  return value && /^[0-9a-f]{6}$/i.test(value) ? `#${value}` : fallback
}

export function previewImage(asset: NativeAsset | undefined): string | undefined {
  if (!asset || !['image/png', 'image/jpeg'].includes(asset.contentType)) return undefined
  const bytes = asset.dataBase64
  // Do not resolve document-controlled remote URLs or embed active SVG content.
  if (!bytes || bytes.length > 4_000_000 || !/^[A-Za-z0-9+/]+={0,2}$/.test(bytes)) return undefined
  return `data:${asset.contentType};base64,${bytes}`
}

export function previewIssue(element: NativeElement, assets: NativeAsset[]): string | undefined {
  const { x, y, cx, cy } = element.transform
  if (![x, y, cx, cy].every(Number.isFinite) || cx <= 0 || cy <= 0) return 'Object preserved; preview geometry unavailable'
  if (element.compatibility.status === 'refused') return 'Unsupported object preserved in the source file'
  if (element.kind === 'group') return 'Grouped content preserved; group preview unavailable'
  if (element.kind === 'table') return 'Table preserved; table preview unavailable'
  if (element.kind === 'chart' && !previewImage(assets.find((asset) => asset.id === element.chart.previewAssetId))) return 'Chart preserved; no embedded preview image'
  if (element.kind === 'picture' && !previewImage(assets.find((asset) => asset.id === element.assetId))) return 'Image preserved; supported preview bytes unavailable'
  if (element.kind === 'shape' && !element.preset) return 'Shape preserved; geometry unavailable'
  return undefined
}

export function previewSlideSize(deck: NativePptxDeck): { width: number; height: number; scale: number } | undefined {
  const { cx, cy } = deck.size
  if (![cx, cy].every((value) => Number.isFinite(value) && value > 0) || cy / cx > 10 || cx / cy > 10) return undefined
  return { width: 960, height: 960 * cy / cx, scale: 960 / cx }
}
