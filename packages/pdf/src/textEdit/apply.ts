import { borrowedFont, loadRequestedFont, type FontRequest, type ResolvedFont, type StandardPdfFont } from './fontResolve.js';
import { resolveObjectMatch, listFlexibleMatches } from './match.js';
import { PdfiumSession } from './pdfium.js';
import { replaceObjectWithLines, replaceObjectsWithSegments, wrapText, type StyledSegment } from './rebuild.js';
import { canReuseTextObject } from './reuse.js';
import { assertEditableText, replaceRange } from './splice.js';
import { readTextObjectStyle, type TextObjectStyle } from './styleRuns.js';

const MAX_OPERATIONS = 1_000;

interface EditBase extends FontRequest {
  readonly page: number;
  readonly oldText: string;
  readonly newText: string;
  readonly occurrence?: number;
}

interface OptionalStyleFields {
  readonly origin?: readonly [number, number];
  readonly newFont?: string;
  readonly newBold?: boolean;
  readonly newColor?: readonly [number, number, number] | readonly [number, number, number, number];
  readonly fontSize?: number;
  readonly lineLeading?: number;
  readonly lineXOffsets?: readonly number[];
  readonly styleRuns?: readonly {
    readonly start: number;
    readonly end: number;
    readonly font?: string;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly color?: readonly [number, number, number] | readonly [number, number, number, number];
    readonly fontSize?: number;
  }[];
}

export interface SurgicalEditSpec extends EditBase {
  readonly kind?: 'surgical';
  readonly rect: readonly [number, number, number, number];
}

export interface BlockEditSpec extends EditBase, OptionalStyleFields {
  readonly kind?: 'block';
  readonly maxWidth?: number;
  readonly rect: readonly [number, number, number, number];
  readonly origin: readonly [number, number];
  readonly fontSize: number;
  readonly lineLeading: number;
  readonly lineHeight?: number;
  readonly align?: 'left' | 'center' | 'right';
}

export interface TextEditSpec extends EditBase, OptionalStyleFields {
  readonly kind?: 'surgical' | 'block';
  readonly rect: readonly [number, number, number, number];
  readonly maxWidth?: number;
  readonly lineHeight?: number;
  readonly align?: 'left' | 'center' | 'right';
}

type InternalEditSpec = TextEditSpec | SurgicalEditSpec | BlockEditSpec;

export interface TextEditFailure {
  readonly index: number;
  readonly page?: number;
  readonly code: 'invalid-spec' | 'not-found' | 'ambiguous-match' | 'cross-object-match' | 'unsupported-layout' | 'unsupported-font' | 'pdfium-error';
  readonly message: string;
  readonly reason?: string;
}

export interface TextEditsResult {
  readonly pdfBytes: Uint8Array;
  readonly bytes: Uint8Array;
  readonly pdf: Uint8Array;
  readonly applied: number;
  readonly skipped: readonly TextEditFailure[];
  readonly failures: readonly TextEditFailure[];
}

export async function applyTextEdits(pdfBytes: Uint8Array, edits: readonly TextEditSpec[]): Promise<TextEditsResult> {
  return applyBatch(pdfBytes, edits);
}

async function applyBatch(pdfBytes: Uint8Array, edits: readonly InternalEditSpec[]): Promise<TextEditsResult> {
  const original = cloneInput(pdfBytes);
  if (!Array.isArray(edits) || edits.length > MAX_OPERATIONS) {
    return result(original, 0, [{ index: 0, code: 'invalid-spec', message: `edits must be an array of at most ${MAX_OPERATIONS} items` }]);
  }
  if (edits.length === 0) return result(original, 0, []);

  const failures = new Map<number, TextEditFailure>();
  edits.forEach((edit, index) => {
    const failure = validateEdit(edit, index)[0];
    if (failure) failures.set(index, failure);
  });
  if (failures.size === edits.length) return result(original, 0, [...failures.values()].sort((left, right) => left.index - right.index));
  const preflightFailures = await preflightEdits(original, edits, failures);
  for (const [index, failure] of preflightFailures) failures.set(index, failure);

  let current = original;
  let applied = 0;
  for (let index = 0; index < edits.length; index += 1) {
    if (failures.has(index)) continue;
    const edit = edits[index]!;
    const session = await PdfiumSession.open(current);
    try {
      const failure = applyOne(session, edit, index);
      if (failure) failures.set(index, failure);
      else {
        current = session.save();
        applied += 1;
      }
    } finally {
      session.close();
    }
  }
  return result(current, applied, [...failures.values()].sort((left, right) => left.index - right.index));
}

async function preflightEdits(
  pdfBytes: Uint8Array,
  edits: readonly InternalEditSpec[],
  knownFailures: ReadonlyMap<number, TextEditFailure>,
): Promise<Map<number, TextEditFailure>> {
  const failures = new Map<number, TextEditFailure>();
  const reserved = new Set<string>();
  const session = await PdfiumSession.open(pdfBytes);
  try {
    for (let index = 0; index < edits.length; index += 1) {
      if (knownFailures.has(index)) continue;
      const edit = edits[index]!;
      try {
        const identities = session.withTextPage(edit.page, (page, textPage) => {
          const resolution = resolveObjectMatch(session.extractText(textPage), edit.oldText, edit.occurrence);
          if (!resolution.ok) throw new MatchFailure(resolution.code, resolution.message);
          return resolution.match.objects.map((object) => `${edit.page}:${session.findObjectIndex(page, object)}`);
        });
        if (identities.some((identity) => identity.endsWith(':-1'))) throw new Error('matched text object is not a direct page object');
        if (identities.some((identity) => reserved.has(identity))) {
          failures.set(index, { index, page: edit.page, code: 'invalid-spec', message: 'edit overlaps another pending text edit' });
          continue;
        }
        identities.forEach((identity) => reserved.add(identity));
      } catch (error) {
        failures.set(index, failureFromError(index, edit.page, error));
      }
    }
  } finally {
    session.close();
  }
  return failures;
}

export function editText(pdfBytes: Uint8Array, spec: readonly SurgicalEditSpec[]): Promise<TextEditsResult> {
  const edits = spec;
  return applyTextEdits(pdfBytes, edits.map((edit) => ({ ...edit, kind: 'surgical' })));
}

export function editBlock(pdfBytes: Uint8Array, spec: readonly BlockEditSpec[]): Promise<TextEditsResult> {
  const edits = spec;
  return applyBatch(pdfBytes, edits.map((edit) => ({ ...edit, kind: 'block' })));
}

function applyOne(session: PdfiumSession, edit: InternalEditSpec, index: number): TextEditFailure | undefined {
  try {
    return session.withPage(edit.page, (page) => {
      const textPage = session.api.FPDFText_LoadPage(page);
      if (!textPage) throw new Error(`PDFium could not create the text index for page ${edit.page}`);
      let selectedText = '';
      let matched: Array<{ object: number; text: string; style: TextObjectStyle; start: number; end: number }> = [];
      try {
        const extracted = session.extractText(textPage);
        const resolution = resolveObjectMatch(extracted, edit.oldText, edit.occurrence);
        if (!resolution.ok) return { index, page: edit.page, code: resolution.code, message: resolution.message };
        selectedText = extracted.text.slice(resolution.match.start, resolution.match.end);
        matched = resolution.match.spans.map((span) => ({
          object: span.object,
          text: session.readObjectText(span.object, textPage),
          style: readTextObjectStyle(session, span.object),
          start: span.start - resolution.match.start,
          end: span.end - resolution.match.start,
        }));
      } finally {
        session.api.FPDFText_ClosePage(textPage);
      }

      if (matched.length > 1) {
        applyMultiObject(session, page, matched, selectedText, edit);
        session.generate(page);
        return undefined;
      }
      const { object, text: objectText, style } = matched[0]!;
      const objectMatches = listFlexibleMatches(objectText, edit.oldText);
      if (objectMatches.length !== 1) {
        return { index, page: edit.page, code: 'ambiguous-match', message: 'the selected page object does not contain one unambiguous copy of the search text' };
      }
      const objectMatch = listFlexibleMatches(objectText, edit.oldText)[0]!;
      const replacement = replaceRange(objectText, objectMatch.start, objectMatch.end, edit.newText);
      const styled = edit as InternalEditSpec & Partial<OptionalStyleFields>;
      if (styled.styleRuns !== undefined) {
        applySegmented(session, page, [object], replacement, objectMatch.start, edit.newText.length, styled, style);
        session.generate(page);
        return undefined;
      }
      const blockRequested = edit.kind === 'block' || styled.lineLeading !== undefined;
      const wholeObject = objectMatch.start === 0 && objectMatch.end === objectText.length;
      const effectiveStyle = wholeObject ? styled : { ...styled, origin: undefined };
      if (!blockRequested && canSegmentForPreservation(objectText, replacement, effectiveStyle, style)) {
        applyPreservingSingle(session, page, object, objectText, replacement, effectiveStyle, style);
      } else if (blockRequested && wholeObject) applyBlock(session, page, object, replacement, edit as BlockEditSpec | TextEditSpec, style);
      else applySurgical(session, page, object, replacement, effectiveStyle, style);
      session.generate(page);
      return undefined;
    });
  } catch (error) {
    return failureFromError(index, edit.page, error);
  }
}

function applyMultiObject(
  session: PdfiumSession,
  page: number,
  sources: readonly { object: number; text: string; style: TextObjectStyle; start: number; end: number }[],
  selectedText: string,
  edit: InternalEditSpec,
): void {
  const styled = edit as InternalEditSpec & Partial<OptionalStyleFields>;
  if (styled.styleRuns !== undefined) {
    applySegmented(session, page, sources.map((source) => source.object), edit.newText, 0, edit.newText.length, styled, sources[0]!.style);
    return;
  }

  const preserved = preservableObjects(sources, selectedText, edit);
  const changed = sources.filter((source) => !preserved.has(source.object));
  const fill = styled.newColor ? normalizeColor(styled.newColor) : undefined;
  for (const source of sources) if (preserved.has(source.object) && fill) session.setFillColor(source.object, fill);
  if (changed.length === 0) return;

  const first = changed[0]!;
  const last = changed.at(-1)!;
  const suffixLength = selectedText.length - last.end;
  let changedText = edit.newText.slice(first.start, edit.newText.length - suffixLength);
  if (first.start > edit.newText.length - suffixLength) {
    preserved.clear();
    changedText = edit.newText;
  }
  const replacedSources = preserved.size ? changed : [...sources];
  const baseSource = replacedSources[0]!;
  const style = overrideStyle(baseSource.style, styled);
  const font = resolveBaseFont(session, changedText, styled, style);
  try {
    replaceObjectsWithSegments(session, page, replacedSources.map((source) => source.object), changedText ? [{ text: changedText, font, style: { ...style, font: font.handle } }] : []);
  } finally {
    font.close();
  }
}

function preservableObjects(
  sources: readonly { object: number; text: string; style: TextObjectStyle; start: number; end: number }[],
  selectedText: string,
  edit: InternalEditSpec,
): Set<number> {
  const styled = edit as InternalEditSpec & Partial<OptionalStyleFields>;
  if (selectedText !== edit.oldText || styled.fontBytes || styled.newFont || styled.newBold !== undefined || styled.fontSize !== undefined) return new Set();
  if (sources.some((source) => Math.abs(source.style.matrix.b) > 0.000_001 || Math.abs(source.style.matrix.c) > 0.000_001)) return new Set();
  if (whitespacePattern(selectedText) !== whitespacePattern(edit.newText)) return new Set();
  const prefix = commonPrefixLength(selectedText, edit.newText);
  const suffix = commonSuffixLength(selectedText, edit.newText, prefix);
  const preserved = new Set<number>();
  for (const source of sources) {
    const direct = selectedText.slice(source.start, source.start + source.text.length) === source.text;
    const trailing = selectedText.slice(source.end - source.text.length, source.end) === source.text;
    const complete = direct || trailing || selectedText.slice(source.start, source.end) === source.text;
    const effectiveStart = trailing ? source.end - source.text.length : source.start;
    const effectiveEnd = direct ? source.start + source.text.length : source.end;
    if (complete && (effectiveEnd <= prefix || effectiveStart >= selectedText.length - suffix)) preserved.add(source.object);
  }
  return preserved;
}

function canSegmentForPreservation(
  oldText: string,
  newText: string,
  edit: InternalEditSpec & Partial<OptionalStyleFields>,
  style: TextObjectStyle,
): boolean {
  if (edit.fontBytes || edit.newFont || edit.newBold !== undefined || edit.fontSize !== undefined || edit.origin) return false;
  if (Math.abs(style.matrix.b) > 0.000_001 || Math.abs(style.matrix.c) > 0.000_001) return false;
  if (whitespacePattern(oldText) !== whitespacePattern(newText)) return false;
  const prefix = commonPrefixLength(oldText, newText);
  const suffix = commonSuffixLength(oldText, newText, prefix);
  return prefix > 0 || suffix > 0;
}

function applyPreservingSingle(
  session: PdfiumSession,
  page: number,
  object: number,
  oldText: string,
  newText: string,
  edit: InternalEditSpec & Partial<OptionalStyleFields>,
  sourceStyle: TextObjectStyle,
): void {
  const prefix = commonPrefixLength(oldText, newText);
  const suffix = commonSuffixLength(oldText, newText, prefix);
  const style = overrideStyle(sourceStyle, edit);
  const font = borrowedFont(sourceStyle.font);
  const textParts = [newText.slice(0, prefix), newText.slice(prefix, newText.length - suffix), newText.slice(newText.length - suffix)];
  replaceObjectsWithSegments(session, page, [object], textParts.filter(Boolean).map((text) => ({ text, font, style: { ...style, font: font.handle } })));
}

function applySegmented(
  session: PdfiumSession,
  page: number,
  sourceObjects: readonly number[],
  replacement: string,
  insertedStart: number,
  insertedLength: number,
  edit: InternalEditSpec & Partial<OptionalStyleFields>,
  sourceStyle: TextObjectStyle,
): void {
  const baseStyle = overrideStyle(sourceStyle, edit);
  const resources: ResolvedFont[] = [];
  const baseFont = resolveBaseFont(session, replacement, edit, baseStyle);
  resources.push(baseFont);
  try {
    const runs = edit.styleRuns ?? [];
    const insertedEnd = insertedStart + insertedLength;
    const boundaries = new Set<number>([0, insertedStart, insertedEnd, replacement.length]);
    for (const run of runs) {
      boundaries.add(insertedStart + run.start);
      boundaries.add(insertedStart + run.end);
    }
    const ordered = [...boundaries].filter((value) => value >= 0 && value <= replacement.length).sort((left, right) => left - right);
    const segments: StyledSegment[] = [];
    for (let index = 0; index + 1 < ordered.length; index += 1) {
      const start = ordered[index]!;
      const end = ordered[index + 1]!;
      if (end <= start) continue;
      const text = replacement.slice(start, end);
      const rawOffset = start - insertedStart;
      const run = rawOffset >= 0 && start < insertedEnd
        ? [...runs].reverse().find((candidate) => candidate.start <= rawOffset && candidate.end > rawOffset)
        : undefined;
      let font = baseFont;
      let style = baseStyle;
      if (run) {
        style = {
          ...baseStyle,
          fill: run.color ? normalizeColor(run.color) : baseStyle.fill,
          fontSize: run.fontSize ?? baseStyle.fontSize,
        };
        if (!Number.isFinite(style.fontSize) || style.fontSize <= 0) throw new RangeError('style run fontSize must be a positive finite number');
        if (run.font || run.bold !== undefined || run.italic !== undefined) {
          font = loadRequestedFont(session, text, { standardFont: styledStandardFont(run.font ?? edit.newFont, run.bold ?? edit.newBold, run.italic) });
          resources.push(font);
        }
      }
      segments.push({ text, font, style: { ...style, font: font.handle } });
    }
    replaceObjectsWithSegments(session, page, sourceObjects, segments);
  } finally {
    for (const resource of resources.reverse()) resource.close();
  }
}

function resolveBaseFont(
  session: PdfiumSession,
  text: string,
  edit: EditBase & Partial<OptionalStyleFields>,
  sourceStyle: TextObjectStyle,
): ResolvedFont {
  if (edit.fontBytes || edit.newFont || edit.newBold !== undefined) {
    return loadRequestedFont(session, text, { ...edit, standardFont: styledStandardFont(edit.newFont, edit.newBold) });
  }
  return borrowedFont(sourceStyle.font);
}

function applySurgical(session: PdfiumSession, page: number, object: number, replacement: string, edit: EditBase & Partial<OptionalStyleFields>, style: TextObjectStyle): void {
  if (replacement.length === 0) {
    if (!session.api.FPDFPage_RemoveObject(page, object)) throw new Error('PDFium could not remove the emptied text object');
    session.api.FPDFPageObj_Destroy(object);
    return;
  }
  if (!edit.fontBytes && !edit.newFont && [...replacement].some((character) => character.codePointAt(0)! > 0xff)) {
    throw new RangeError('no available font can draw replacement text within the bounded v1 policy');
  }
  const requestedStyle = overrideStyle(style, edit);
  const requiresRebuild = Boolean(edit.fontBytes || edit.newFont || edit.newBold !== undefined || edit.fontSize !== undefined || edit.newColor !== undefined || edit.origin !== undefined);
  if (!requiresRebuild) {
    const decision = canReuseTextObject(style, replacement);
    if (!decision.reusable) throw new RangeError(decision.reason);
    if (!session.assignText(object, replacement)) throw new RangeError('source font cannot encode the replacement text');
    return;
  }
  const font = edit.fontBytes || edit.newFont || edit.newBold !== undefined
    ? loadRequestedFont(session, replacement, { ...edit, standardFont: legacyStandardFont(edit.newFont, edit.newBold) })
    : borrowedFont(style.font);
  try {
    replaceObjectWithLines(session, page, object, [{ text: replacement, estimatedWidth: 0 }], font, requestedStyle, requestedStyle.fontSize * 1.2, 'left', 0);
  } finally { font.close(); }
}

function applyBlock(session: PdfiumSession, page: number, object: number, replacement: string, edit: BlockEditSpec | TextEditSpec, style: TextObjectStyle): void {
  style = overrideStyle(style, edit);
  if (Math.abs(style.matrix.b) > 0.000_001 || Math.abs(style.matrix.c) > 0.000_001) throw new RangeError('rotated or sheared block text is outside the v1 layout boundary');
  const maxWidth = edit.maxWidth ?? (edit.rect ? Math.abs(edit.rect[2] - edit.rect[0]) : undefined);
  if (!maxWidth) throw new RangeError('block editing requires maxWidth or a non-zero rect width');
  const lines = wrapText(replacement, maxWidth, style.fontSize);
  const lineHeight = edit.lineHeight ?? edit.lineLeading ?? style.fontSize * 1.2;
  if (!Number.isFinite(lineHeight) || lineHeight <= 0) throw new RangeError('lineHeight must be a positive finite number');
  const font = edit.fontBytes || edit.newFont || edit.newBold !== undefined
    ? loadRequestedFont(session, replacement.replace(/\n/g, ''), { ...edit, standardFont: legacyStandardFont(edit.newFont, edit.newBold) })
    : borrowedFont(style.font);
  try { replaceObjectWithLines(session, page, object, lines, font, style, lineHeight, edit.align ?? 'left', maxWidth, edit.lineXOffsets); }
  finally { font.close(); }
}

function validateEdit(edit: InternalEditSpec, index: number): TextEditFailure[] {
  try {
    if (!edit || typeof edit !== 'object') throw new TypeError('edit must be an object');
    assertEditableText(edit.oldText, 'oldText');
    assertEditableText(edit.newText, 'newText', true);
    if (!Number.isSafeInteger(edit.page) || edit.page < 1) throw new RangeError('page must be a positive integer');
    if (edit.occurrence !== undefined && (!Number.isSafeInteger(edit.occurrence) || edit.occurrence < 0)) throw new RangeError('occurrence must be a non-negative integer');
    if (edit.kind === 'block' && edit.maxWidth !== undefined && (!Number.isFinite(edit.maxWidth) || edit.maxWidth <= 0)) throw new RangeError('maxWidth must be a positive finite number');
    const styled = edit as InternalEditSpec & Partial<OptionalStyleFields>;
    if (styled.styleRuns !== undefined) {
      if (!Array.isArray(styled.styleRuns)) throw new TypeError('styleRuns must be an array');
      for (const run of styled.styleRuns) {
        if (!Number.isSafeInteger(run.start) || !Number.isSafeInteger(run.end) || run.start < 0 || run.end <= run.start || run.end > edit.newText.length) {
          throw new RangeError('style run offsets must be a non-empty half-open range within newText');
        }
        if (run.color) normalizeColor(run.color);
        if (run.fontSize !== undefined && (!Number.isFinite(run.fontSize) || run.fontSize <= 0)) throw new RangeError('style run fontSize must be positive');
      }
    }
    return [];
  } catch (error) {
    return [{ index, page: edit?.page, code: 'invalid-spec', message: errorMessage(error) }];
  }
}

function cloneInput(bytes: Uint8Array): Uint8Array {
  if (!(bytes instanceof Uint8Array)) throw new TypeError('pdfBytes must be a Uint8Array');
  return bytes;
}

function result(pdfBytes: Uint8Array, applied: number, failures: readonly TextEditFailure[]): TextEditsResult {
  const skipped = failures.map((failure) => ({ ...failure, reason: failure.message }));
  return { pdfBytes, bytes: pdfBytes, pdf: pdfBytes, applied, skipped, failures };
}

function overrideStyle(style: TextObjectStyle, edit: EditBase & Partial<OptionalStyleFields>): TextObjectStyle {
  const fill = edit.newColor ? normalizeColor(edit.newColor) : style.fill;
  const matrix = edit.origin ? { ...style.matrix, e: edit.origin[0], f: edit.origin[1] } : style.matrix;
  const fontSize = edit.fontSize ?? style.fontSize;
  if (!Number.isFinite(fontSize) || fontSize <= 0) throw new RangeError('fontSize must be a positive finite number');
  return { ...style, fill, matrix, fontSize };
}

function normalizeColor(color: NonNullable<OptionalStyleFields['newColor']>): readonly [number, number, number, number] {
  const values = [color[0], color[1], color[2], color.length === 4 ? color[3] : 255];
  if (!values.every((value) => Number.isInteger(value) && value! >= 0 && value! <= 255)) throw new RangeError('newColor channels must be integers from 0 to 255');
  return values as [number, number, number, number];
}

function classifyFailure(error: unknown): TextEditFailure['code'] {
  const message = errorMessage(error);
  if (/font|glyph|encode|sfnt|TrueType/i.test(message)) return 'unsupported-font';
  if (/rotated|sheared|layout|lineHeight|maxWidth/i.test(message)) return 'unsupported-layout';
  return 'pdfium-error';
}

function failureFromError(index: number, page: number | undefined, error: unknown): TextEditFailure {
  if (error instanceof MatchFailure) return { index, page, code: error.code, message: error.message };
  return { index, page, code: classifyFailure(error), message: errorMessage(error) };
}

class MatchFailure extends Error {
  readonly code: Extract<TextEditFailure['code'], 'not-found' | 'ambiguous-match' | 'cross-object-match'>;
  constructor(code: MatchFailure['code'], message: string) {
    super(message);
    this.code = code;
  }
}

function errorMessage(error: unknown): string { return error instanceof Error ? error.message : String(error); }

function styledStandardFont(name?: string, bold = false, italic = false): StandardPdfFont {
  const family = name?.toLowerCase();
  if (family === 'times' || family === 'times-roman') {
    if (bold && italic) return 'Times-BoldItalic';
    if (bold) return 'Times-Bold';
    if (italic) return 'Times-Italic';
    return 'Times-Roman';
  }
  if (family === 'courier') {
    if (bold && italic) return 'Courier-BoldOblique';
    if (bold) return 'Courier-Bold';
    if (italic) return 'Courier-Oblique';
    return 'Courier';
  }
  if (family === 'symbol') return 'Symbol';
  if (family === 'zapfdingbats') return 'ZapfDingbats';
  if (bold && italic) return 'Helvetica-BoldOblique';
  if (bold) return 'Helvetica-Bold';
  if (italic) return 'Helvetica-Oblique';
  return 'Helvetica';
}

const legacyStandardFont = styledStandardFont;

function whitespacePattern(value: string): string {
  return value.match(/\s+/gu)?.join('\0') ?? '';
}

function commonPrefixLength(left: string, right: string): number {
  let index = 0;
  while (index < left.length && index < right.length && left[index] === right[index]) index += 1;
  return index;
}

function commonSuffixLength(left: string, right: string, prefixLength: number): number {
  let length = 0;
  while (length < left.length - prefixLength && length < right.length - prefixLength
    && left[left.length - 1 - length] === right[right.length - 1 - length]) length += 1;
  return length;
}
