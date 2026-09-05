export const MAX_TEXT_UNITS = 1_048_576;

export function assertEditableText(value: string, label: string, allowEmpty = false): void {
  if (typeof value !== 'string') throw new TypeError(`${label} must be a string`);
  if (!allowEmpty && value.length === 0) throw new RangeError(`${label} must not be empty`);
  if (value.length > MAX_TEXT_UNITS) throw new RangeError(`${label} exceeds ${MAX_TEXT_UNITS} UTF-16 units`);
  if (value.includes('\0')) throw new RangeError(`${label} must not contain U+0000`);
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next < 0xdc00 || next > 0xdfff) throw new RangeError(`${label} contains an unpaired high surrogate`);
      index += 1;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      throw new RangeError(`${label} contains an unpaired low surrogate`);
    }
  }
}

export function replaceRange(source: string, start: number, end: number, replacement: string): string {
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start || end > source.length) {
    throw new RangeError('replacement range is outside the source text');
  }
  if (start > 0 && isLow(source.charCodeAt(start))) throw new RangeError('replacement starts inside a surrogate pair');
  if (end < source.length && isLow(source.charCodeAt(end))) throw new RangeError('replacement ends inside a surrogate pair');
  assertEditableText(replacement, 'replacement', true);
  return source.slice(0, start) + replacement + source.slice(end);
}

export function replaceExact(source: string, needle: string, replacement: string, occurrence = 0): string {
  assertEditableText(needle, 'search text');
  assertEditableText(replacement, 'replacement', true);
  if (!Number.isSafeInteger(occurrence) || occurrence < 0) throw new RangeError('occurrence must be a non-negative integer');
  let from = 0;
  for (let seen = 0; ; seen += 1) {
    const at = source.indexOf(needle, from);
    if (at < 0) throw new RangeError(`search occurrence ${occurrence} was not found`);
    if (seen === occurrence) return replaceRange(source, at, at + needle.length, replacement);
    from = at + needle.length;
  }
}

function isLow(unit: number): boolean {
  return unit >= 0xdc00 && unit <= 0xdfff;
}

export function foldMap(text: string, foldInput: ((character: string) => string) | boolean = (character) => character.normalize('NFKC')): any {
  const fold = typeof foldInput === 'function' ? foldInput : (character: string) => foldInput ? character.normalize('NFKC') : character;
  let folded = '';
  const units: string[] = [];
  const starts: number[] = [];
  const ends: number[] = [];
  let sourceOffset = 0;
  for (const character of text) {
    if (/\s/u.test(character)) {
      if (foldInput === true && units.at(-1) !== ' ') {
        units.push(' '); starts.push(sourceOffset); ends.push(sourceOffset + character.length); folded += ' ';
      } else if (foldInput === true) {
        ends[ends.length - 1] = sourceOffset + character.length;
      }
      sourceOffset += character.length;
      continue;
    }
    const value = fold(character);
    folded += value;
    for (const unit of value) { units.push(unit); starts.push(sourceOffset); ends.push(sourceOffset + character.length); }
    sourceOffset += character.length;
  }
  return { text: folded, sourceOffsets: [...starts, text.length], units, idx: [...starts, text.length], end: ends };
}

export function indexOfUnits(haystack: readonly (number | string)[] | string, needle: readonly (number | string)[] | string, from = 0): number {
  const source = typeof haystack === 'string' ? [...haystack] : [...haystack];
  const target = typeof needle === 'string' ? [...needle] : [...needle];
  if (target.length === 0) return -1;
  outer: for (let index = Math.max(0, from); index <= source.length - target.length; index += 1) {
    for (let offset = 0; offset < target.length; offset += 1) if (source[index + offset] !== target[offset]) continue outer;
    return index;
  }
  return -1;
}

export function mergeEngineCodepoints(...parts: any[]): any {
  const [engineText, oldText, newText] = parts.map((part) => String(part ?? ''));
  if (oldText === newText) return engineText;
  const engine = foldMap(engineText);
  const old = foldMap(oldText);
  const next = foldMap(newText);
  if (engine.units.length !== old.units.length) return newText;
  let prefix = 0;
  while (prefix < old.units.length && prefix < next.units.length && old.units[prefix] === next.units[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < old.units.length - prefix && suffix < next.units.length - prefix
    && old.units[old.units.length - 1 - suffix] === next.units[next.units.length - 1 - suffix]) suffix += 1;
  const enginePrefixEnd = prefix === 0 ? 0 : engine.end[prefix - 1];
  const engineSuffixStart = suffix === 0 ? engineText.length : engine.idx[engine.units.length - suffix];
  const newMiddleStart = prefix === 0 ? 0 : next.end[prefix - 1];
  const newMiddleEnd = suffix === 0 ? newText.length : next.idx[next.units.length - suffix];
  return engineText.slice(0, enginePrefixEnd) + newText.slice(newMiddleStart, newMiddleEnd) + engineText.slice(engineSuffixStart);
}

export function spliceIntoEngine(...args: any[]): any {
  if (args.length === 3) return mergeEngineCodepoints(args[0], args[1], args[2]);
  const [source, start, deleteCount, replacement = ''] = args as [string, number, number, string?];
  if (!Number.isSafeInteger(deleteCount) || deleteCount < 0) throw new RangeError('deleteCount must be a non-negative integer');
  return replaceRange(source, start, start + deleteCount, replacement);
}
