declare module 'bidi-js' {
  interface EmbeddingLevels {
    levels: Uint8Array
    paragraphs: Array<{ start: number; end: number; level: number }>
  }

  interface BidiApi {
    getEmbeddingLevels(text: string, direction?: 'ltr' | 'rtl' | 'auto'): EmbeddingLevels
    getBidiCharTypeName(character: string): string
  }

  export default function bidiFactory(): BidiApi
}
