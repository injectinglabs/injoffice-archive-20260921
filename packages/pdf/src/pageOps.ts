import { PDFDocument, degrees, type PDFPage } from 'pdf-lib';
import type {
  CropOp,
  DeleteOp,
  InsertBlankOp,
  NUpOp,
  PageOpSpec,
  PageRange,
  PageSelector,
  PdfDocumentInfo,
  ReorderOp,
  ResizeOp,
  RotateOp,
} from './types.js';

/** Reads page count and per-page geometry without altering the document. */
export async function readInfo(bytes: Uint8Array): Promise<PdfDocumentInfo> {
  const doc = await PDFDocument.load(bytes);
  const pages = doc.getPages().map((page, i) => {
    const { width, height } = page.getSize();
    return { index: i + 1, width, height, rotation: normalizeRotation(page.getRotation().angle) };
  });
  return { pageCount: pages.length, pages };
}

const DEFAULT_PAGE_SIZE = { width: 612, height: 792 }; // US Letter, points

export function normalizeRotation(angle: number): 0 | 90 | 180 | 270 {
  const normalized = ((Math.round(angle / 90) * 90) % 360 + 360) % 360;
  return normalized as 0 | 90 | 180 | 270;
}

/** Resolves a 1-indexed page selector against a page count. Throws on out-of-range indices. */
export function resolveSelector(selector: PageSelector, pageCount: number): number[] {
  const pages = selector === 'all' ? Array.from({ length: pageCount }, (_, i) => i + 1) : [...selector];
  for (const p of pages) {
    if (!Number.isInteger(p) || p < 1 || p > pageCount) {
      throw new Error(`page ${p} is out of range for a ${pageCount}-page document`);
    }
  }
  return pages;
}

async function applyRotate(bytes: Uint8Array, op: RotateOp): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const pages = resolveSelector(op.pages, doc.getPageCount());
  for (const p of pages) {
    const page = doc.getPage(p - 1);
    const current = page.getRotation().angle;
    page.setRotation(degrees(normalizeRotation(current + op.degrees)));
  }
  return doc.save();
}

async function applyDelete(bytes: Uint8Array, op: DeleteOp): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const pageCount = doc.getPageCount();
  const toRemove = resolveSelector(op.pages, pageCount);
  if (toRemove.length >= pageCount) {
    throw new Error('cannot delete every page in a document');
  }
  // Remove highest index first so earlier indices stay valid.
  const descending = [...new Set(toRemove)].sort((a, b) => b - a);
  for (const p of descending) doc.removePage(p - 1);
  return doc.save();
}

async function applyReorder(bytes: Uint8Array, op: ReorderOp): Promise<Uint8Array> {
  const src = await PDFDocument.load(bytes);
  const pageCount = src.getPageCount();
  const isPermutation =
    op.order.length === pageCount && new Set(op.order).size === pageCount && op.order.every((p) => p >= 1 && p <= pageCount);
  if (!isPermutation) {
    throw new Error(`reorder "order" must be a permutation of 1..${pageCount}`);
  }
  const out = await PDFDocument.create();
  const copied = await out.copyPages(
    src,
    op.order.map((p) => p - 1),
  );
  copied.forEach((page) => out.addPage(page));
  return out.save();
}

async function applyInsertBlank(bytes: Uint8Array, op: InsertBlankOp): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const pageCount = doc.getPageCount();
  if (!Number.isInteger(op.at) || op.at < 1 || op.at > pageCount + 1) {
    throw new Error(`insertBlank "at" ${op.at} is out of range for a ${pageCount}-page document (1..${pageCount + 1})`);
  }
  const size = op.size ?? (pageCount > 0 ? doc.getPage(0).getSize() : DEFAULT_PAGE_SIZE);
  doc.insertPage(op.at - 1, [size.width, size.height]);
  return doc.save();
}

function computeCropOrResizeTarget(page: PDFPage, box: CropOp['box']) {
  page.setCropBox(box.x, box.y, box.width, box.height);
  page.setMediaBox(box.x, box.y, box.width, box.height);
}

async function applyCrop(bytes: Uint8Array, op: CropOp): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const pages = resolveSelector(op.pages, doc.getPageCount());
  for (const p of pages) computeCropOrResizeTarget(doc.getPage(p - 1), op.box);
  return doc.save();
}

async function applyResize(bytes: Uint8Array, op: ResizeOp): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes);
  const pages = resolveSelector(op.pages, doc.getPageCount());
  for (const p of pages) {
    const page = doc.getPage(p - 1);
    const { width: origW, height: origH } = page.getSize();
    if (op.fit === 'stretch') {
      page.scale(op.width / origW, op.height / origH);
      page.setSize(op.width, op.height);
    } else {
      const factor = Math.min(op.width / origW, op.height / origH);
      page.scale(factor, factor);
      const scaledW = origW * factor;
      const scaledH = origH * factor;
      const offsetX = (op.width - scaledW) / 2;
      const offsetY = (op.height - scaledH) / 2;
      page.setMediaBox(-offsetX, -offsetY, op.width, op.height);
      page.setCropBox(-offsetX, -offsetY, op.width, op.height);
    }
  }
  return doc.save();
}

async function applyNUp(bytes: Uint8Array, op: NUpOp): Promise<Uint8Array> {
  const src = await PDFDocument.load(bytes);
  const sheet = op.pageSize ?? DEFAULT_PAGE_SIZE;
  const { cols, rows } = nUpGrid(op.n);
  const cellW = sheet.width / cols;
  const cellH = sheet.height / rows;

  const out = await PDFDocument.create();
  const srcPages = src.getPages();
  const embedded = await out.embedPages(srcPages);

  for (let sheetStart = 0; sheetStart < embedded.length; sheetStart += op.n) {
    const page = out.addPage([sheet.width, sheet.height]);
    const group = embedded.slice(sheetStart, sheetStart + op.n);
    group.forEach((ep, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const scale = Math.min(cellW / ep.width, cellH / ep.height);
      const drawW = ep.width * scale;
      const drawH = ep.height * scale;
      const cellX = col * cellW;
      const cellY = sheet.height - (row + 1) * cellH;
      page.drawPage(ep, {
        x: cellX + (cellW - drawW) / 2,
        y: cellY + (cellH - drawH) / 2,
        width: drawW,
        height: drawH,
      });
    });
  }
  return out.save();
}

function nUpGrid(n: NUpOp['n']): { cols: number; rows: number } {
  switch (n) {
    case 2:
      return { cols: 1, rows: 2 };
    case 4:
      return { cols: 2, rows: 2 };
    case 6:
      return { cols: 2, rows: 3 };
    case 9:
      return { cols: 3, rows: 3 };
  }
}

/** Applies a sequence of page operations, each a fresh fail-closed transform over the previous bytes. */
export async function applyPageOps(bytes: Uint8Array, ops: PageOpSpec[]): Promise<Uint8Array> {
  let current = bytes;
  for (const op of ops) {
    current = await applyOne(current, op);
  }
  return current;
}

async function applyOne(bytes: Uint8Array, op: PageOpSpec): Promise<Uint8Array> {
  switch (op.type) {
    case 'rotate':
      return applyRotate(bytes, op);
    case 'delete':
      return applyDelete(bytes, op);
    case 'reorder':
      return applyReorder(bytes, op);
    case 'insertBlank':
      return applyInsertBlank(bytes, op);
    case 'crop':
      return applyCrop(bytes, op);
    case 'resize':
      return applyResize(bytes, op);
    case 'nUp':
      return applyNUp(bytes, op);
  }
}

/** Splits a document into one output per inclusive 1-indexed page range. */
export async function splitDocument(bytes: Uint8Array, ranges: PageRange[]): Promise<Uint8Array[]> {
  const src = await PDFDocument.load(bytes);
  const pageCount = src.getPageCount();
  const outputs: Uint8Array[] = [];
  for (const range of ranges) {
    if (range.from < 1 || range.to > pageCount || range.from > range.to) {
      throw new Error(`invalid split range ${range.from}-${range.to} for a ${pageCount}-page document`);
    }
    const indices = Array.from({ length: range.to - range.from + 1 }, (_, i) => range.from - 1 + i);
    const out = await PDFDocument.create();
    const copied = await out.copyPages(src, indices);
    copied.forEach((page) => out.addPage(page));
    outputs.push(await out.save());
  }
  return outputs;
}

/** Concatenates whole documents, in order, into one output. */
export async function mergeDocuments(documents: Uint8Array[]): Promise<Uint8Array> {
  if (documents.length === 0) throw new Error('mergeDocuments requires at least one document');
  const out = await PDFDocument.create();
  for (const bytes of documents) {
    const src = await PDFDocument.load(bytes);
    const copied = await out.copyPages(src, src.getPages().map((_, i) => i));
    copied.forEach((page) => out.addPage(page));
  }
  return out.save();
}
