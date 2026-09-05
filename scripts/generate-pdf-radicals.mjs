import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const sourcePath = resolve(root, 'scripts/ucd/EquivalentUnifiedIdeograph-17.0.0.txt');
const outputPath = resolve(root, 'packages/pdf/src/textEdit/radicals.ts');
const PINNED_SHA256 = '38619c05a17e771554000fe604afee92e10eb49e0616ecf0c87af3c1eb0f4320';

const tsChar = (cp) => (cp > 0xffff ? `'\\u{${cp.toString(16)}}'` : `'\\u${cp.toString(16).padStart(4, '0')}'`);

const bytes = readFileSync(sourcePath);
const sha256 = createHash('sha256').update(bytes).digest('hex');
if (sha256 !== PINNED_SHA256) {
  throw new Error(`EquivalentUnifiedIdeograph-17.0.0.txt SHA-256 ${sha256} does not match pinned ${PINNED_SHA256}`);
}

const pairs = [];
for (const raw of bytes.toString('utf8').split(/\r?\n/)) {
  const line = raw.replace(/#.*/, '').trim();
  if (!line) continue;
  const [rangeField, unifiedField] = line.split(';').map((part) => part.trim());
  if (!rangeField || !unifiedField) continue;
  const [startText, endText = startText] = rangeField.split('..');
  const start = Number.parseInt(startText, 16);
  const end = Number.parseInt(endText, 16);
  const unified = Number.parseInt(unifiedField, 16);
  if (![start, end, unified].every((n) => Number.isSafeInteger(n) && n >= 0 && n <= 0x10ffff) || end < start) {
    throw new Error(`invalid EquivalentUnifiedIdeograph mapping: ${line}`);
  }
  for (let cp = start; cp <= end; cp++) pairs.push([cp, unified]);
}

if (pairs.length === 0) throw new Error('no EquivalentUnifiedIdeograph mappings parsed');

const entries = pairs.map(([from, to]) => `  ${tsChar(from)}: ${tsChar(to)},`).join('\n');

writeFileSync(
  outputPath,
  `/** CJK radical and stroke → equivalent unified ideograph.
 * Generated from Unicode 17.0.0 EquivalentUnifiedIdeograph.txt
 * (https://www.unicode.org/Public/17.0.0/ucd/EquivalentUnifiedIdeograph.txt,
 * SHA-256 ${sha256}). Unicode License v3; see /LICENSE-UNICODE.txt.
 *
 * pdf.js may extract Kangxi / Radicals Supplement / CJK Stroke codepoints
 * from a font cmap while PDFium returns the unified ideograph. NFKC only
 * decomposes Kangxi (U+2F00–U+2FD5); this table covers the rest.
 */
export const RADICAL_EQUIV: Record<string, string> = {
${entries}
};

/** Fold radical-block and CJK-stroke codepoints to their unified ideographs. */
export function foldRadicals(s: string): string {
  let out = '';
  for (const ch of s) out += RADICAL_EQUIV[ch] ?? ch;
  return out;
}
`,
);
console.log(`wrote ${pairs.length} mappings to ${outputPath}`);
