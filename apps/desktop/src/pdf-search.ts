export interface PdfSearchHit { start: number; end: number; spans: number[] }
export interface PdfSearchResult { hits: Array<{ page: number; match: number }>; pagesScanned: number; totalPages: number; limited: boolean }

export function findPdfTextMatches(parts: readonly string[], query: string): PdfSearchHit[] {
  const needle = query.trim().toLocaleLowerCase();
  if (!needle || needle.length > 500) return [];
  const folded = parts.map(part => part.toLocaleLowerCase());
  const haystack = folded.join(' ');
  const ranges: Array<{ start: number; end: number }> = [];
  let cursor = 0;
  for (const part of folded) { ranges.push({ start: cursor, end: cursor + part.length }); cursor += part.length + 1; }
  const found: PdfSearchHit[] = [];
  for (let at = haystack.indexOf(needle); at >= 0 && found.length < 1000; at = haystack.indexOf(needle, at + Math.max(1, needle.length))) {
    found.push({ start: at, end: at + needle.length, spans: ranges.flatMap((range, index) => range.start < at + needle.length && range.end > at ? [index] : []) });
  }
  return found;
}

/** Sequential extraction keeps only result coordinates, with explicit work/result limits. */
export async function searchPdfDocument(totalPages: number, readPage: (page: number) => Promise<string[]>, query: string, signal: AbortSignal, progress: (page: number) => void = () => {}): Promise<PdfSearchResult> {
  if (!Number.isInteger(totalPages) || totalPages < 1 || !query.trim() || query.length > 500) throw new Error('Enter text to find in this PDF.');
  const result: PdfSearchResult = { hits: [], pagesScanned: 0, totalPages, limited: false };
  let characters = 0;
  for (let page = 1; page <= Math.min(totalPages, 2000); page++) {
    signal.throwIfAborted();
    const parts = await readPage(page);
    signal.throwIfAborted();
    characters += parts.reduce((sum, part) => sum + part.length, 0);
    if (characters > 10000000) { result.limited = true; break; }
    const matches = findPdfTextMatches(parts, query);
    for (let match = 0; match < matches.length && result.hits.length < 10000; match++) result.hits.push({ page, match });
    result.pagesScanned = page;
    progress(page);
    if (result.hits.length >= 10000 || matches.length >= 1000) { result.limited = true; break; }
  }
  if (result.pagesScanned < totalPages) result.limited = true;
  return result;
}
