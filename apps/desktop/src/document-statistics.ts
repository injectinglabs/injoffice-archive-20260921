import type { NativeDocxDocumentV1, NativeDocxParagraphV1, NativeDocxRunV1 } from '../../../packages/docs/src/nativeContract';

const segmenter = new Intl.Segmenter(undefined, { granularity: 'word' });

function count(text: string) {
  let words = 0, characters = 0;
  for (const part of segmenter.segment(text)) if (part.isWordLike) words++;
  for (const _ of text) characters++;
  return { words, characters };
}

function runText(run: NativeDocxRunV1) {
  return run.kind === 'text' && !run.properties?.hidden ? (run.text ?? '') : '';
}

function paragraphText(paragraph: NativeDocxParagraphV1, draft?: { runId: string; text: string }) {
  return paragraph.runs.map(run => (run.kind === 'text' && !run.properties?.hidden && run.id === draft?.runId ? draft.text : runText(run))).join('');
}

/** Paragraph totals prevent an ordinary keystroke from recounting the entire document. */
export function documentStatistics(document: NativeDocxDocumentV1) {
  const paragraphs = document.body.blocks.flatMap(block => block.paragraph ? [block.paragraph] : block.table?.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs)) ?? []);
  const entries = new Map(paragraphs.map(paragraph => [paragraph.id, { paragraph, ...count(paragraphText(paragraph)) }]));
  let words = 0, characters = 0;
  for (const entry of entries.values()) { words += entry.words; characters += entry.characters; }
  return { entries, words, characters, paragraphs: paragraphs.length };
}

export function statisticsWithDraft(base: ReturnType<typeof documentStatistics>, draft?: { paragraphId: string; runId: string; text: string }) {
  const entry = draft && base.entries.get(draft.paragraphId), changed = entry && count(paragraphText(entry.paragraph, draft));
  return { words: base.words + (changed && entry ? changed.words - entry.words : 0), characters: base.characters + (changed && entry ? changed.characters - entry.characters : 0), paragraphs: base.paragraphs };
}
