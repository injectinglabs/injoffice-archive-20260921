/** Local, strict UTF-8 CSV/TSV. No type inference: fields remain literal text. */
export type DelimitedFormat = 'csv' | 'tsv';
export const DELIMITED_LIMITS = { bytes: 4 * 1024 * 1024, exportBytes: 16 * 1024 * 1024, rows: 10000, columns: 256, fields: 50000, fieldLength: 32767 } as const;
const encoder = new TextEncoder();
function separator(format: DelimitedFormat): string { if (format !== 'csv' && format !== 'tsv') throw new Error('Choose CSV or TSV.'); return format === 'csv' ? ',' : '\t'; }
export function parseDelimited(bytes: Uint8Array, format: DelimitedFormat): string[][] {
  const delimiter = separator(format);
  if (!(bytes instanceof Uint8Array) || bytes.byteLength > DELIMITED_LIMITS.bytes) throw new Error('Import at most 4 MiB of UTF-8 text.');
  let text: string;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch { throw new Error('The file must use valid UTF-8 encoding.'); }
  if (text.includes('\0')) throw new Error('NUL bytes are not supported in delimited text.');
  if (!text.length) return [];
  const rows: string[][] = []; let row: string[] = [], field = '', state: 'start' | 'plain' | 'quoted' | 'closed' = 'start', fields = 0, endedRecord = false;
  const append = (value: string) => { field += value; if (field.length > DELIMITED_LIMITS.fieldLength) throw new Error('A field exceeds 32,767 text units.'); };
  const finishField = () => { if (++fields > DELIMITED_LIMITS.fields || row.length >= DELIMITED_LIMITS.columns) throw new Error('Import at most 50,000 fields and 256 columns.'); row.push(field); field = ''; state = 'start'; };
  const finishRow = () => { finishField(); if (rows.length >= DELIMITED_LIMITS.rows) throw new Error('Import at most 10,000 rows.'); rows.push(row); row = []; endedRecord = true; };
  for (let i = 0; i < text.length; i++) {
    const c = text[i]; endedRecord = false;
    if (state === 'quoted') {
      if (c === '"') { if (text[i + 1] === '"') { append('"'); i++; } else state = 'closed'; }
      else append(c);
      continue;
    }
    if (c === delimiter) { finishField(); continue; }
    if (c === '\r' || c === '\n') { if (c === '\r' && text[i + 1] === '\n') i++; finishRow(); continue; }
    if (state === 'closed') throw new Error(`Unexpected text after a closing quote at character ${i + 1}.`);
    if (c === '"') { if (state !== 'start') throw new Error(`Quote inside an unquoted field at character ${i + 1}.`); state = 'quoted'; }
    else { state = 'plain'; append(c); }
  }
  if (state === 'quoted') throw new Error('The final quoted field is not closed.');
  if (!endedRecord) finishRow();
  return rows;
}
export function encodeDelimited(rows: readonly (readonly string[])[], format: DelimitedFormat): Uint8Array {
  const delimiter = separator(format);
  if (rows.length > DELIMITED_LIMITS.rows) throw new Error('Export at most 10,000 rows.');
  let fields = 0, size = 0; const lines: string[] = [];
  for (const row of rows) {
    if (row.length === 0) throw new Error('Each exported record needs at least one field.');
    if (row.length > DELIMITED_LIMITS.columns || (fields += row.length) > DELIMITED_LIMITS.fields) throw new Error('Export at most 50,000 fields and 256 columns.');
    const line = row.map(value => { if (typeof value !== 'string' || value.length > DELIMITED_LIMITS.fieldLength || value.includes('\0') || /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value)) throw new Error('A field contains unsupported text or exceeds 32,767 text units.'); return value.includes(delimiter) || /["\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value; }).join(delimiter) + '\r\n';
    size += encoder.encode(line).byteLength; if (size > DELIMITED_LIMITS.exportBytes) throw new Error('Export at most 16 MiB of UTF-8 text.'); lines.push(line);
  }
  return encoder.encode(lines.join(''));
}
