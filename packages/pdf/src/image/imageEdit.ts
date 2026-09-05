import { PNG } from 'pngjs'
import { PDFDocument } from 'pdf-lib'
import type { WrappedPdfiumModule } from '@embedpdf/pdfium'
import type { ImageEditSpec, ImageEditsResult, ImageLayer, PageImageRef } from './types.js'
import { allocate, exclusivePdfium, memory, release, saveDocument, withDocument } from './pdfiumBridge.js'

type Rect = readonly [number, number, number, number]

const PAGE_OBJECT_TEXT = 1
const PAGE_OBJECT_IMAGE = 3
const BITMAP_BGRA = 4
const RECT_TOLERANCE = 0.5
const MAX_IMAGE_PIXELS = 64_000_000

function canonicalRect(rect: Rect): [number, number, number, number] {
  if (rect.length !== 4 || rect.some((value) => !Number.isFinite(value))) {
    throw new TypeError('rectangle must contain four finite numbers')
  }
  return [Math.min(rect[0], rect[2]), Math.min(rect[1], rect[3]), Math.max(rect[0], rect[2]), Math.max(rect[1], rect[3])]
}

function sameRect(left: Rect, right: Rect): boolean {
  const a = canonicalRect(left)
  const b = canonicalRect(right)
  return a.every((value, index) => Math.abs(value - b[index]) <= RECT_TOLERANCE)
}

function normalizeTurns(value = 0): number {
  if (!Number.isInteger(value) || value < 0 || value > 3) throw new TypeError('quarterTurns must be an integer from 0 through 3')
  return value
}

function counterTurns(rotation = 0): number {
  const normalized = ((rotation % 360) + 360) % 360
  if (![0, 90, 180, 270].includes(normalized)) throw new TypeError('rotate must resolve to 0, 90, 180, or 270 degrees')
  return ((4 - normalized / 90) % 4)
}

function matrixFor(rectInput: Rect, turns: number): [number, number, number, number, number, number] {
  const [x1, y1, x2, y2] = canonicalRect(rectInput)
  const width = x2 - x1
  const height = y2 - y1
  if (width <= 0 || height <= 0) throw new TypeError('image rectangle must have positive width and height')
  if (turns === 1) return [0, -height, width, 0, x1, y2]
  if (turns === 2) return [-width, 0, 0, -height, x2, y2]
  if (turns === 3) return [0, height, -width, 0, x2, y1]
  return [width, 0, 0, height, x1, y1]
}

function objectBounds(module: WrappedPdfiumModule, object: number): [number, number, number, number] | undefined {
  const ptr = allocate(module, 16)
  try {
    if (!module.FPDFPageObj_GetBounds(object, ptr, ptr + 4, ptr + 8, ptr + 12)) return undefined
    const offset = ptr / 4
    return [
      memory(module).HEAPF32[offset],
      memory(module).HEAPF32[offset + 1],
      memory(module).HEAPF32[offset + 2],
      memory(module).HEAPF32[offset + 3],
    ]
  } finally {
    release(module, ptr)
  }
}

function firstTextIndex(module: WrappedPdfiumModule, page: number): number {
  const count = module.FPDFPage_CountObjects(page)
  for (let index = 0; index < count; index += 1) {
    if (module.FPDFPageObj_GetType(module.FPDFPage_GetObject(page, index)) === PAGE_OBJECT_TEXT) return index
  }
  return -1
}

function insertAtLayer(module: WrappedPdfiumModule, page: number, object: number, layer: ImageLayer): void {
  if (layer === 'aboveText') {
    module.FPDFPage_InsertObject(page, object)
    return
  }
  const firstText = firstTextIndex(module, page)
  const index = firstText < 0 ? 0 : firstText
  if (!module.FPDFPage_InsertObjectAtIndex(page, object, index)) {
    throw new Error('PDFium could not insert image below text')
  }
}

function insertAtIndex(module: WrappedPdfiumModule, page: number, object: number, index: number): void {
  const bounded = Math.max(0, Math.min(index, module.FPDFPage_CountObjects(page)))
  if (!module.FPDFPage_InsertObjectAtIndex(page, object, bounded)) throw new Error('PDFium could not restore image paint order')
}

function findImage(module: WrappedPdfiumModule, page: number, rect: Rect): { object: number; index: number } | undefined {
  const count = module.FPDFPage_CountObjects(page)
  for (let index = 0; index < count; index += 1) {
    const object = module.FPDFPage_GetObject(page, index)
    if (module.FPDFPageObj_GetType(object) !== PAGE_OBJECT_IMAGE) continue
    const bounds = objectBounds(module, object)
    if (bounds && sameRect(bounds, rect)) return { object, index }
  }
  return undefined
}

function createImageObject(
  module: WrappedPdfiumModule,
  document: number,
  page: number,
  imageBase64: string,
  rect: Rect,
  turns: number,
): number {
  let decoded: ReturnType<typeof PNG.sync.read>
  try {
    decoded = PNG.sync.read(Buffer.from(imageBase64, 'base64'), { skipRescale: true })
  } catch {
    throw new Error('could not decode PNG image data')
  }
  if (decoded.width <= 0 || decoded.height <= 0 || decoded.width * decoded.height > MAX_IMAGE_PIXELS) {
    throw new Error('PNG dimensions are outside the supported range')
  }
  const stride = decoded.width * 4
  const pixels = allocate(module, stride * decoded.height)
  const imageObject = module.FPDFPageObj_NewImageObj(document)
  if (!imageObject) {
    release(module, pixels)
    throw new Error('PDFium could not create an image object')
  }
  let bitmap = 0
  try {
    const heap = memory(module).HEAPU8
    for (let index = 0; index < decoded.data.length; index += 4) {
      heap[pixels + index] = decoded.data[index + 2]
      heap[pixels + index + 1] = decoded.data[index + 1]
      heap[pixels + index + 2] = decoded.data[index]
      heap[pixels + index + 3] = decoded.data[index + 3]
    }
    bitmap = module.FPDFBitmap_CreateEx(decoded.width, decoded.height, BITMAP_BGRA, pixels, stride)
    if (!bitmap) throw new Error('PDFium could not create a bitmap')
    const pages = allocate(module, 4)
    try {
      memory(module).HEAP32[pages / 4] = page
      if (!module.FPDFImageObj_SetBitmap(pages, 1, imageObject, bitmap)) throw new Error('PDFium could not attach PNG pixels')
    } finally {
      release(module, pages)
    }
    if (!module.FPDFImageObj_SetMatrix(imageObject, ...matrixFor(rect, turns))) throw new Error('PDFium could not position image')
    return imageObject
  } catch (error) {
    module.FPDFPageObj_Destroy(imageObject)
    throw error
  } finally {
    if (bitmap) module.FPDFBitmap_Destroy(bitmap)
    release(module, pixels)
  }
}

function pageFailure(page: number, pageCount: number): string | undefined {
  return Number.isInteger(page) && page >= 1 && page <= pageCount ? undefined : 'page does not exist'
}

export async function applyImageEdits(bytes: Uint8Array, edits: ImageEditSpec[]): Promise<ImageEditsResult> {
  if (edits.length === 0) return { bytes, skipped: [] }
  return exclusivePdfium((module) => withDocument(module, bytes, (document) => {
    const pageCount = module.FPDF_GetPageCount(document)
    const skipped: ImageEditsResult['skipped'] = []
    let changed = false

    for (let editIndex = 0; editIndex < edits.length; editIndex += 1) {
      const edit = edits[editIndex]
      const invalidPage = pageFailure(edit.page, pageCount)
      if (invalidPage) {
        skipped.push({ editIndex, page: edit.page, reason: invalidPage })
        continue
      }
      const page = module.FPDF_LoadPage(document, edit.page - 1)
      if (!page) {
        skipped.push({ editIndex, page: edit.page, reason: 'PDFium could not load page' })
        continue
      }
      let editChanged = false
      try {
        if (edit.kind === 'insertImage') {
          const image = createImageObject(module, document, page, edit.image, edit.rect, counterTurns(edit.rotate))
          try {
            insertAtLayer(module, page, image, edit.layer)
          } catch (error) {
            module.FPDFPageObj_Destroy(image)
            throw error
          }
          editChanged = true
        } else {
          const match = findImage(module, page, edit.oldRect)
          if (!match) {
            skipped.push({ editIndex, page: edit.page, reason: 'image could not be located at oldRect' })
            continue
          }
          if (edit.kind === 'deleteImage') {
            if (!module.FPDFPage_RemoveObject(page, match.object)) throw new Error('PDFium could not remove image')
            module.FPDFPageObj_Destroy(match.object)
            editChanged = true
          } else if (edit.kind === 'transformImage') {
            if (!module.FPDFImageObj_SetMatrix(match.object, ...matrixFor(edit.rect, normalizeTurns(edit.quarterTurns)))) {
              throw new Error('PDFium could not transform image')
            }
            if (edit.layer) {
              if (!module.FPDFPage_RemoveObject(page, match.object)) throw new Error('PDFium could not move image layer')
              insertAtLayer(module, page, match.object, edit.layer)
            }
            editChanged = true
          } else {
            const replacement = createImageObject(module, document, page, edit.image, edit.rect, normalizeTurns(edit.quarterTurns))
            try {
              if (!module.FPDFPage_RemoveObject(page, match.object)) throw new Error('PDFium could not remove replaced image')
              if (edit.layer) insertAtLayer(module, page, replacement, edit.layer)
              else insertAtIndex(module, page, replacement, match.index)
              module.FPDFPageObj_Destroy(match.object)
            } catch (error) {
              module.FPDFPageObj_Destroy(replacement)
              throw error
            }
            editChanged = true
          }
        }
        if (editChanged && !module.FPDFPage_GenerateContent(page)) throw new Error('PDFium could not regenerate page content')
        changed ||= editChanged
      } catch (error) {
        skipped.push({ editIndex, page: edit.page, reason: error instanceof Error ? error.message : String(error) })
      } finally {
        module.FPDF_ClosePage(page)
      }
    }
    return { bytes: changed ? saveDocument(module, document) : bytes, skipped }
  }))
}

export async function listPageImages(bytes: Uint8Array): Promise<PageImageRef[]> {
  return exclusivePdfium((module) => withDocument(module, bytes, (document) => {
    const output: PageImageRef[] = []
    const pageCount = module.FPDF_GetPageCount(document)
    for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
      const page = module.FPDF_LoadPage(document, pageIndex)
      if (!page) continue
      try {
        const textIndex = firstTextIndex(module, page)
        const count = module.FPDFPage_CountObjects(page)
        for (let objectIndex = 0; objectIndex < count; objectIndex += 1) {
          const object = module.FPDFPage_GetObject(page, objectIndex)
          if (module.FPDFPageObj_GetType(object) !== PAGE_OBJECT_IMAGE) continue
          const rect = objectBounds(module, object)
          if (rect) output.push({ page: pageIndex + 1, rect: canonicalRect(rect), aboveText: textIndex >= 0 && objectIndex > textIndex })
        }
      } finally {
        module.FPDF_ClosePage(page)
      }
    }
    return output
  }))
}

export async function verifyImageEdits(bytes: Uint8Array, edits: { page: number; rect: Rect }[]): Promise<{ page: number; reason: string }[]> {
  const pageCount = (await PDFDocument.load(bytes)).getPageCount()
  const images = await listPageImages(bytes)
  return edits.flatMap((edit) => {
    if (!Number.isInteger(edit.page) || edit.page < 1 || edit.page > pageCount) return [{ page: edit.page, reason: 'page missing from saved output' }]
    return images.some((image) => image.page === edit.page && sameRect(image.rect, edit.rect))
      ? []
      : [{ page: edit.page, reason: 'image missing from saved output' }]
  })
}
