export interface ExtractedUnit {
  readonly start: number;
  readonly end: number;
  readonly object: number;
}

export interface ExtractedText {
  readonly text: string;
  readonly units: readonly ExtractedUnit[];
}

export interface ObjectMatch {
  readonly start: number;
  readonly end: number;
  readonly object: number;
  readonly objects: readonly number[];
  readonly spans: readonly { object: number; start: number; end: number }[];
}

export type PageTextObj = Record<string, any>;

export type MatchResolution =
  | { readonly ok: true; readonly match: ObjectMatch }
  | { readonly ok: false; readonly code: 'not-found' | 'ambiguous-match' | 'cross-object-match'; readonly message: string };

export function listExactMatches(text: string, needle: string): ReadonlyArray<{ start: number; end: number }> {
  if (needle.length === 0) return [];
  const matches: Array<{ start: number; end: number }> = [];
  let cursor = 0;
  while (cursor <= text.length - needle.length) {
    const start = text.indexOf(needle, cursor);
    if (start < 0) break;
    matches.push({ start, end: start + needle.length });
    cursor = start + needle.length;
  }
  return matches;
}

export function listFlexibleMatches(text: string, needle: string): ReadonlyArray<{ start: number; end: number }> {
  const exact = listExactMatches(text, needle);
  if (exact.length || !/\s/u.test(needle)) return exact;
  const parts = needle.trim().split(/\s+/u).map(escapeRegExp);
  if (parts.length === 0 || parts.some((part) => part.length === 0)) return [];
  const pattern = new RegExp(parts.join('\\s+'), 'gu');
  return [...text.matchAll(pattern)].map((match) => ({ start: match.index, end: match.index + match[0].length }));
}

export function resolveObjectMatch(extracted: ExtractedText, needle: string, occurrence?: number): MatchResolution {
  const candidates = listFlexibleMatches(extracted.text, needle);
  if (candidates.length === 0) return { ok: false, code: 'not-found', message: 'search text could not be located on the page' };
  if (occurrence === undefined && candidates.length !== 1) {
    return { ok: false, code: 'ambiguous-match', message: `search text has ${candidates.length} non-overlapping matches; specify occurrence` };
  }
  const selected = candidates[occurrence ?? 0];
  if (!selected) return { ok: false, code: 'not-found', message: `search occurrence ${occurrence} was not found` };
  const covered = extracted.units.filter((unit) => unit.end > selected.start && unit.start < selected.end);
  if (covered.length === 0 || covered[0]!.start !== selected.start || covered.at(-1)!.end !== selected.end) {
    return { ok: false, code: 'cross-object-match', message: 'match does not align to complete extracted Unicode units' };
  }
  const objects: number[] = [];
  const spans: Array<{ object: number; start: number; end: number }> = [];
  for (const unit of covered) {
    if (unit.object === 0) {
      if (!/^\s$/u.test(extracted.text.slice(unit.start, unit.end))) {
        return { ok: false, code: 'cross-object-match', message: 'match includes generated non-whitespace text' };
      }
      continue;
    }
    const previous = spans.at(-1);
    if (previous?.object === unit.object) previous.end = unit.end;
    else {
      if (objects.includes(unit.object)) return { ok: false, code: 'cross-object-match', message: 'text object appears in disjoint reading-order spans' };
      objects.push(unit.object);
      spans.push({ object: unit.object, start: unit.start, end: unit.end });
    }
  }
  if (objects.length === 0) return { ok: false, code: 'cross-object-match', message: 'match contains no editable PDF text object' };
  return { ok: true, match: { ...selected, object: objects[0]!, objects, spans } };
}

export function canReuseFont(...args: any[]): boolean {
  if (typeof args[0] === 'string' && Array.isArray(args[1]) && Array.isArray(args[2])) {
    const [replacement, matches, all] = args as [string, PageTextObj[], PageTextObj[]];
    if (matches.length !== 1) return false;
    const font = matches[0]?.font ?? matches[0]?.fontName ?? matches[0]?.fontId;
    const available = all.filter((item) => (item.font ?? item.fontName ?? item.fontId) === font).map((item) => String(item.text ?? '')).join('');
    return [...replacement].filter((character) => !/\s/u.test(character)).every((character) => available.includes(character));
  }
  const [object, replacement = ''] = args as [PageTextObj, string?];
  const supported = object.supportedText ?? object.text ?? object.chars;
  if (typeof supported === 'string') return [...replacement].every((character) => supported.includes(character));
  if (supported && typeof supported[Symbol.iterator] === 'function') {
    const set = new Set(supported as Iterable<unknown>);
    return [...replacement].every((character) => set.has(character) || set.has(character.codePointAt(0)));
  }
  return false;
}

export function resolveMatch(...args: any[]): any {
  const [objects, rect, oldText, newText, useMerge = false] = args as [PageTextObj[], [number, number, number, number], string, string, boolean?];
  if (!Array.isArray(objects) || !Array.isArray(rect) || typeof oldText !== 'string') return { reason: 'invalid match request' };
  const touching = objects.filter((object) => intersects(rect, object.rect ?? object.bounds ?? [0, 0, 0, 0]));
  const ordered = [...touching].sort((left, right) => Number(left.index ?? left.objectIndex ?? 0) - Number(right.index ?? right.objectIndex ?? 0));
  const joined = ordered.map((object) => String(object.text ?? '')).join(' ');
  const compact = (value: string) => value.replace(/\s+/gu, ' ').trim();
  if (compact(joined) === compact(oldText)) return { matches: ordered, newText: String(newText ?? '').trim(), whole: true };
  const containing = ordered.filter((object) => String(object.text ?? '').includes(oldText));
  if (containing.length === 1 && !useMerge) {
    const object = containing[0]!;
    const whole = compact(String(object.text)) === compact(oldText);
    return { matches: [object], newText: whole ? String(newText ?? '').trim() : String(object.text).replace(oldText, String(newText ?? '').trim()), whole };
  }
  return { reason: useMerge && containing.length ? 'fragment reflow is unsupported' : 'text could not be located in the edit rectangle' };
}

function intersects(left: readonly number[], right: readonly number[]): boolean {
  return left.length === 4 && right.length === 4 && left[0]! <= right[2]! + 2 && left[2]! >= right[0]! - 2 && left[1]! <= right[3]! + 2 && left[3]! >= right[1]! - 2;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&');
}
