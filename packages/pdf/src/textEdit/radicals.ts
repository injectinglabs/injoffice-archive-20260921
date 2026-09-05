export function unicodeFallbackCandidates(character: string): readonly string[] {
  if ([...character].length !== 1) return [];
  const candidates = new Set<string>();
  for (const form of ['NFC', 'NFKC'] as const) {
    const normalized = character.normalize(form);
    if (normalized !== character && [...normalized].length === 1) candidates.add(normalized);
  }
  return [...candidates];
}

export const RADICAL_EQUIV: Record<string, string> = buildRadicalEquivalents();

export function foldRadicals(text: string): string {
  return [...text].map((character) => RADICAL_EQUIV[character] ?? unicodeFallbackCandidates(character)[0] ?? character).join('');
}

function buildRadicalEquivalents(): Record<string, string> {
  const equivalents: Record<string, string> = { '\u2e85': '\u4ebb' };
  for (let scalar = 0x2f00; scalar <= 0x2fd5; scalar += 1) {
    const radical = String.fromCodePoint(scalar);
    const unified = radical.normalize('NFKC');
    if (unified !== radical) equivalents[radical] = unified;
  }
  return Object.freeze(equivalents);
}
