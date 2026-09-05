import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { init, type WrappedPdfiumModule } from '@embedpdf/pdfium';
import type { ExtractedText, ExtractedUnit } from './match.js';
import type { TextMatrix, TextObjectStyle } from './styleRuns.js';

export const MAX_PDF_BYTES = 128 * 1024 * 1024;
export const FPDF_PAGEOBJ_TEXT = 1;
export const FPDF_PAGEOBJ_PATH = 2;
export const FPDF_PAGEOBJ_FORM = 5;
export type Pdfium = any;
const MAX_PAGE_CHARACTERS = 10_000_000;
const MAX_PAGE_OBJECTS = 1_000_000;

let modulePromise: Promise<WrappedPdfiumModule> | undefined;

export async function getPdfium(): Promise<WrappedPdfiumModule> {
  modulePromise ??= initializePdfium();
  return modulePromise;
}

export async function loadPdfium(): Promise<Pdfium> {
  return (await getPdfium()).pdfium;
}

let operationTail: Promise<void> = Promise.resolve();

export function chainPdfium<T>(operation: (pdfium: Pdfium) => T | Promise<T>): Promise<T> {
  const run = operationTail.then(async () => operation(await loadPdfium()));
  operationTail = run.then(() => undefined, () => undefined);
  return run;
}

export function withDocument<T>(pdfium: Pdfium, pdfBytes: Uint8Array, operation: (document: number) => T): T {
  if (!(pdfBytes instanceof Uint8Array) || pdfBytes.byteLength === 0 || pdfBytes.byteLength > MAX_PDF_BYTES) {
    throw new RangeError(`pdfBytes must contain 1..${MAX_PDF_BYTES} bytes`);
  }
  const pointer = pdfium._malloc(pdfBytes.byteLength);
  if (!pointer) throw new Error('PDFium could not allocate source memory');
  pdfium.HEAPU8.set(pdfBytes, pointer);
  const document = pdfium._FPDF_LoadMemDocument64(pointer, pdfBytes.byteLength, 0);
  if (!document) {
    pdfium._free(pointer);
    throw new Error(`PDFium could not open the document (error ${pdfium._FPDF_GetLastError()})`);
  }
  try { return operation(document); }
  finally { pdfium._FPDF_CloseDocument(document); pdfium._free(pointer); }
}

export function saveDoc(pdfium: Pdfium, document: number): Uint8Array {
  const writer = pdfium._PDFiumExt_OpenFileWriter();
  if (!writer) throw new Error('PDFium could not create an output writer');
  try {
    if (!pdfium._PDFiumExt_SaveAsCopy(document, writer)) throw new Error('PDFium could not serialize the document');
    const size = pdfium._PDFiumExt_GetFileWriterSize(writer);
    if (!Number.isSafeInteger(size) || size <= 0 || size > MAX_PDF_BYTES * 2) throw new RangeError('PDFium produced an invalid output size');
    const pointer = pdfium._malloc(size);
    if (!pointer) throw new Error('PDFium could not allocate output memory');
    try {
      if (pdfium._PDFiumExt_GetFileWriterData(writer, pointer, size) !== size) throw new Error('PDFium returned incomplete output');
      return pdfium.HEAPU8.slice(pointer, pointer + size);
    } finally { pdfium._free(pointer); }
  } finally { pdfium._PDFiumExt_CloseFileWriter(writer); }
}

async function initializePdfium(): Promise<WrappedPdfiumModule> {
  const resolved = import.meta.resolve('@embedpdf/pdfium/pdfium.wasm');
  const file = await readFile(fileURLToPath(resolved));
  const wasmBinary = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
  const api = await init({ wasmBinary });
  api.PDFiumExt_Init();
  return api;
}

export class PdfiumSession {
  readonly api: WrappedPdfiumModule;
  readonly document: number;
  readonly pageCount: number;
  readonly #sourcePointer: number;
  #closed = false;

  private constructor(api: WrappedPdfiumModule, sourcePointer: number, document: number) {
    this.api = api;
    this.#sourcePointer = sourcePointer;
    this.document = document;
    this.pageCount = api.FPDF_GetPageCount(document);
  }

  static async open(pdfBytes: Uint8Array): Promise<PdfiumSession> {
    if (!(pdfBytes instanceof Uint8Array)) throw new TypeError('pdfBytes must be a Uint8Array');
    if (pdfBytes.byteLength === 0 || pdfBytes.byteLength > MAX_PDF_BYTES) {
      throw new RangeError(`pdfBytes must contain 1..${MAX_PDF_BYTES} bytes`);
    }
    const api = await getPdfium();
    const pointer = api.pdfium.wasmExports.malloc(pdfBytes.byteLength);
    if (!pointer) throw new Error('PDFium could not allocate source memory');
    heap(api).set(pdfBytes, pointer);
    const document = api.FPDF_LoadMemDocument64(pointer, pdfBytes.byteLength, '');
    if (!document) {
      const code = api.FPDF_GetLastError();
      api.pdfium.wasmExports.free(pointer);
      throw new Error(`PDFium could not open the document (error ${code})`);
    }
    return new PdfiumSession(api, pointer, document);
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    this.api.FPDF_CloseDocument(this.document);
    this.api.pdfium.wasmExports.free(this.#sourcePointer);
  }

  withPage<T>(pageNumber: number, operation: (page: number) => T): T {
    this.assertPage(pageNumber);
    const page = this.api.FPDF_LoadPage(this.document, pageNumber - 1);
    if (!page) throw new Error(`PDFium could not load page ${pageNumber}`);
    try { return operation(page); } finally { this.api.FPDF_ClosePage(page); }
  }

  withTextPage<T>(pageNumber: number, operation: (page: number, textPage: number) => T): T {
    return this.withPage(pageNumber, (page) => {
      const textPage = this.api.FPDFText_LoadPage(page);
      if (!textPage) throw new Error(`PDFium could not create the text index for page ${pageNumber}`);
      try { return operation(page, textPage); } finally { this.api.FPDFText_ClosePage(textPage); }
    });
  }

  extractText(textPage: number): ExtractedText {
    const count = this.api.FPDFText_CountChars(textPage);
    if (count < 0 || count > MAX_PAGE_CHARACTERS) throw new RangeError('PDFium reported an invalid or excessive page character count');
    let text = '';
    const units: ExtractedUnit[] = [];
    for (let index = 0; index < count; index += 1) {
      const scalar = this.api.FPDFText_GetUnicode(textPage, index);
      const character = validScalar(scalar) ? String.fromCodePoint(scalar) : '\ufffd';
      const start = text.length;
      text += character;
      units.push({ start, end: text.length, object: this.api.FPDFText_GetTextObject(textPage, index) });
    }
    return { text, units };
  }

  readObjectText(object: number, textPage: number): string {
    const byteLength = this.api.FPDFTextObj_GetText(object, textPage, 0, 0);
    if (byteLength < 2 || byteLength > (MAX_PAGE_CHARACTERS + 1) * 2) throw new RangeError('PDFium reported an invalid text-object length');
    return this.withAllocation(byteLength, (pointer) => {
      const written = this.api.FPDFTextObj_GetText(object, textPage, pointer, byteLength);
      if (written !== byteLength) throw new Error('PDFium did not return the complete text object');
      return this.api.pdfium.UTF16ToString(pointer);
    });
  }

  assignText(object: number, text: string): boolean {
    if (text.length === 0) return false;
    return this.withWideString(text, (pointer) => this.api.FPDFText_SetText(object, pointer));
  }

  readTextObjectStyle(object: number): TextObjectStyle {
    const fontSize = this.withAllocation(4, (pointer) => {
      if (!this.api.FPDFTextObj_GetFontSize(object, pointer)) throw new Error('could not read source font size');
      return this.api.pdfium.getValue(pointer, 'float') as number;
    });
    const matrix = this.withAllocation(24, (pointer) => {
      if (!this.api.FPDFPageObj_GetMatrix(object, pointer)) throw new Error('could not read source text matrix');
      return {
        a: this.api.pdfium.getValue(pointer, 'float') as number,
        b: this.api.pdfium.getValue(pointer + 4, 'float') as number,
        c: this.api.pdfium.getValue(pointer + 8, 'float') as number,
        d: this.api.pdfium.getValue(pointer + 12, 'float') as number,
        e: this.api.pdfium.getValue(pointer + 16, 'float') as number,
        f: this.api.pdfium.getValue(pointer + 20, 'float') as number,
      };
    });
    const fill = this.withAllocation(16, (pointer) => {
      if (!this.api.FPDFPageObj_GetFillColor(object, pointer, pointer + 4, pointer + 8, pointer + 12)) {
        throw new Error('could not read source fill color');
      }
      return [0, 4, 8, 12].map((offset) => this.api.pdfium.getValue(pointer + offset, 'i32') as number) as unknown as readonly [number, number, number, number];
    });
    const font = this.api.FPDFTextObj_GetFont(object);
    if (!font) throw new Error('could not read source font');
    const renderMode = this.api.FPDFTextObj_GetTextRenderMode(object);
    if (renderMode < 0) throw new Error('could not read source text render mode');
    return { font, fontSize, fill, renderMode, matrix };
  }

  setTextObjectStyle(object: number, style: TextObjectStyle, lineOffset = 0): void {
    this.setFillColor(object, style.fill);
    if (!this.api.FPDFTextObj_SetTextRenderMode(object, style.renderMode)) throw new Error('could not set text render mode');
    const matrix: TextMatrix = { ...style.matrix, e: style.matrix.e, f: style.matrix.f - lineOffset };
    this.setMatrix(object, matrix);
  }

  setFillColor(object: number, fill: readonly [number, number, number, number]): void {
    const [red, green, blue, alpha] = fill;
    if (!this.api.FPDFPageObj_SetFillColor(object, red, green, blue, alpha)) throw new Error('could not set text fill color');
  }

  readObjectBounds(object: number): readonly [number, number, number, number] {
    return this.withAllocation(16, (pointer) => {
      if (!this.api.FPDFPageObj_GetBounds(object, pointer, pointer + 4, pointer + 8, pointer + 12)) throw new Error('could not read text object bounds');
      return [0, 4, 8, 12].map((offset) => this.api.pdfium.getValue(pointer + offset, 'float') as number) as unknown as readonly [number, number, number, number];
    });
  }

  setMatrix(object: number, matrix: TextMatrix): void {
    this.withAllocation(24, (pointer) => {
      const values = [matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f];
      values.forEach((value, index) => this.api.pdfium.setValue(pointer + index * 4, value, 'float'));
      if (!this.api.FPDFPageObj_SetMatrix(object, pointer)) throw new Error('could not set text matrix');
    });
  }

  findObjectIndex(page: number, object: number): number {
    const count = this.api.FPDFPage_CountObjects(page);
    if (count < 0 || count > MAX_PAGE_OBJECTS) throw new RangeError('PDFium reported an invalid or excessive page-object count');
    for (let index = 0; index < count; index += 1) if (this.api.FPDFPage_GetObject(page, index) === object) return index;
    return -1;
  }

  generate(page: number): void {
    if (!this.api.FPDFPage_GenerateContent(page)) throw new Error('PDFium could not regenerate page content');
  }

  save(): Uint8Array {
    const writer = this.api.PDFiumExt_OpenFileWriter();
    if (!writer) throw new Error('PDFium could not create an output writer');
    try {
      const result = this.api.PDFiumExt_SaveAsCopy(this.document, writer);
      if (!result) throw new Error('PDFium could not serialize the edited document');
      const size = this.api.PDFiumExt_GetFileWriterSize(writer);
      if (!Number.isSafeInteger(size) || size <= 0 || size > MAX_PDF_BYTES * 2) throw new RangeError('PDFium produced an invalid or excessive output size');
      return this.withAllocation(size, (pointer) => {
        const copied = this.api.PDFiumExt_GetFileWriterData(writer, pointer, size);
        if (copied !== size) throw new Error('PDFium did not return the complete serialized document');
        return heap(this.api).slice(pointer, pointer + size);
      });
    } finally {
      this.api.PDFiumExt_CloseFileWriter(writer);
    }
  }

  withAllocation<T>(size: number, operation: (pointer: number) => T): T {
    const pointer = this.api.pdfium.wasmExports.malloc(size);
    if (!pointer) throw new Error(`PDFium could not allocate ${size} bytes`);
    try { return operation(pointer); } finally { this.api.pdfium.wasmExports.free(pointer); }
  }

  withBytes<T>(bytes: Uint8Array, operation: (pointer: number) => T): T {
    return this.withAllocation(bytes.byteLength, (pointer) => {
      heap(this.api).set(bytes, pointer);
      return operation(pointer);
    });
  }

  withWideString<T>(text: string, operation: (pointer: number) => T): T {
    const size = (text.length + 1) * 2;
    return this.withAllocation(size, (pointer) => {
      this.api.pdfium.stringToUTF16(text, pointer, size);
      return operation(pointer);
    });
  }

  private assertPage(pageNumber: number): void {
    if (!Number.isSafeInteger(pageNumber) || pageNumber < 1 || pageNumber > this.pageCount) {
      throw new RangeError(`page ${pageNumber} does not exist; valid pages are 1 through ${this.pageCount}`);
    }
  }
}

function validScalar(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff && !(value >= 0xd800 && value <= 0xdfff);
}

function heap(pdfium: WrappedPdfiumModule): Uint8Array {
  return (pdfium.pdfium as typeof pdfium.pdfium & { HEAPU8: Uint8Array }).HEAPU8;
}
