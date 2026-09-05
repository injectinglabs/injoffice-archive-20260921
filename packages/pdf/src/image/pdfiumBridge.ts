import { readFile } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { init, type WrappedPdfiumModule } from '@embedpdf/pdfium'

let modulePromise: Promise<WrappedPdfiumModule> | undefined
let operationTail: Promise<void> = Promise.resolve()

export interface PdfiumMemory {
  HEAPU8: Uint8Array
  HEAP32: Int32Array
  HEAPF32: Float32Array
}

export function memory(module: WrappedPdfiumModule): PdfiumMemory {
  return module.pdfium as unknown as PdfiumMemory
}

async function loadModule(): Promise<WrappedPdfiumModule> {
  if (!modulePromise) {
    modulePromise = (async () => {
      const wasmPath = createRequire(import.meta.url).resolve('@embedpdf/pdfium/pdfium.wasm')
      const wasmBinary = await readFile(wasmPath)
      const module = await init({ wasmBinary })
      module.PDFiumExt_Init()
      return module
    })()
  }
  return modulePromise
}

export function allocate(module: WrappedPdfiumModule, size: number): number {
  const ptr = module.pdfium.wasmExports.malloc(Math.max(1, size))
  if (!ptr) throw new Error(`PDFium allocation failed (${size} bytes)`)
  return ptr
}

export function release(module: WrappedPdfiumModule, ptr: number): void {
  if (ptr) module.pdfium.wasmExports.free(ptr)
}

export async function exclusivePdfium<T>(action: (module: WrappedPdfiumModule) => Promise<T> | T): Promise<T> {
  let unlock!: () => void
  const predecessor = operationTail
  operationTail = new Promise<void>((resolve) => { unlock = resolve })
  await predecessor
  try {
    return await action(await loadModule())
  } finally {
    unlock()
  }
}

export function withDocument<T>(
  module: WrappedPdfiumModule,
  bytes: Uint8Array,
  action: (document: number) => Promise<T> | T,
): Promise<T> | T {
  const source = allocate(module, bytes.byteLength)
  memory(module).HEAPU8.set(bytes, source)
  const document = module.FPDF_LoadMemDocument64(source, bytes.byteLength, '')
  if (!document) {
    release(module, source)
    throw new Error(`PDFium could not load document (error ${module.FPDF_GetLastError()})`)
  }
  const close = () => {
    module.FPDF_CloseDocument(document)
    release(module, source)
  }
  try {
    const result = action(document)
    if (result instanceof Promise) return result.finally(close)
    close()
    return result
  } catch (error) {
    close()
    throw error
  }
}

export function saveDocument(module: WrappedPdfiumModule, document: number): Uint8Array {
  const writer = module.PDFiumExt_OpenFileWriter()
  if (!writer) throw new Error('PDFium could not create an output writer')
  try {
    if (!module.PDFiumExt_SaveAsCopy(document, writer)) throw new Error('PDFium could not serialize the document')
    const size = module.PDFiumExt_GetFileWriterSize(writer)
    if (size <= 0) throw new Error('PDFium produced an empty document')
    const output = allocate(module, size)
    try {
      const copied = module.PDFiumExt_GetFileWriterData(writer, output, size)
      if (copied <= 0) throw new Error('PDFium could not read serialized bytes')
      return memory(module).HEAPU8.slice(output, output + size)
    } finally {
      release(module, output)
    }
  } finally {
    module.PDFiumExt_CloseFileWriter(writer)
  }
}
