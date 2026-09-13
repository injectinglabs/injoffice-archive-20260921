/** Unicode17 scalar-array adapter over the attributed bidi-js algorithm. */
export interface EmbeddingLevels { levels: Uint8Array; paragraphs: { start: number; end: number; level: number }[] }
export function getEmbeddingLevels(scalars: readonly string[], baseDirection?: 'ltr' | 'rtl' | 'auto'): EmbeddingLevels
export function getReorderedIndices(scalars: readonly string[], result: EmbeddingLevels, start?: number, end?: number): number[]
export function getBidiCharTypeName(scalar: string): string
