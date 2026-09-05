/** A single page's geometry, 1-indexed to match how agents and UI refer to pages. */
export interface PdfPageInfo {
  index: number;
  width: number;
  height: number;
  /** Normalized to one of 0/90/180/270. */
  rotation: 0 | 90 | 180 | 270;
}

export interface PdfDocumentInfo {
  pageCount: number;
  pages: PdfPageInfo[];
}

export type PageSelector = number[] | 'all';

export interface RotateOp {
  type: 'rotate';
  pages: PageSelector;
  /** Added to each selected page's current rotation, then normalized mod 360. */
  degrees: 90 | 180 | 270 | -90 | -180 | -270;
}

export interface DeleteOp {
  type: 'delete';
  pages: number[];
}

export interface ReorderOp {
  type: 'reorder';
  /** A full permutation of the document's current 1-indexed page numbers. */
  order: number[];
}

export interface InsertBlankOp {
  type: 'insertBlank';
  /** 1-indexed position the new page takes after insertion. */
  at: number;
  size?: { width: number; height: number };
}

export interface CropOp {
  type: 'crop';
  pages: PageSelector;
  /** PDF points, origin at the page's bottom-left (pdf-lib/PDF convention). */
  box: { x: number; y: number; width: number; height: number };
}

export interface ResizeOp {
  type: 'resize';
  pages: PageSelector;
  width: number;
  height: number;
  /** 'stretch' scales x/y independently to fill the target exactly; 'contain' preserves aspect ratio and centers. */
  fit: 'stretch' | 'contain';
}

export interface NUpOp {
  type: 'nUp';
  n: 2 | 4 | 6 | 9;
  pageSize?: { width: number; height: number };
}

export type PageOpSpec = RotateOp | DeleteOp | ReorderOp | InsertBlankOp | CropOp | ResizeOp | NUpOp;

export interface PageRange {
  /** Inclusive, 1-indexed. */
  from: number;
  to: number;
}
