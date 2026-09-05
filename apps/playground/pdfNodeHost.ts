import type { IncomingMessage, ServerResponse } from 'node:http'
import { applyTextEdits, applyTextInserts } from '../../packages/pdf/src/textEdit/index.ts'
import { applyImageEdits, listPageImages } from '../../packages/pdf/src/image/index.ts'
import { applyOcrLayer, renderPageToPng } from '../../packages/pdf/src/ocr/index.ts'
import { listRedactableFormFields, redactFormFields } from '../../packages/pdf/src/annotate/formRedact.ts'
import { applyWholeTextRedactions } from '../../packages/pdf/src/redaction/apply.ts'
import { applyWholeImageRedactions } from '../../packages/pdf/src/redaction/image.ts'
import { applyWholeAnnotationRedactions, listWholeAnnotationRedactionTargets } from '../../packages/pdf/src/redaction/annotation.ts'
import { applyWholePathRedactions, listPagePaths } from '../../packages/pdf/src/redaction/path.ts'
import { applyWholeXObjectRedactions, listPageXObjects } from '../../packages/pdf/src/redaction/xobject.ts'
import { PDF_NODE_PREFIX } from './src/pdfHostRoutes.ts'

type Json = Record<string, unknown>

const STAMP_PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='

function send(res: ServerResponse, status: number, body: Json) {
  res.statusCode = status
  res.setHeader('Content-Type', 'application/json')
  res.end(JSON.stringify(body))
}

function bytesFrom(raw: unknown): Uint8Array {
  if (typeof raw !== 'string' || raw.length === 0) throw new Error('bytes must be a base64 string')
  return Uint8Array.from(Buffer.from(raw, 'base64'))
}

function toB64(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64')
}

function numberTuple(value: unknown, fallback: readonly [number, number, number, number]): [number, number, number, number] {
  if (!Array.isArray(value) || value.length !== 4) return [...fallback]
  const next = value.map((item) => Number(item))
  if (next.some((item) => !Number.isFinite(item))) return [...fallback]
  return next as [number, number, number, number]
}

async function readJson(req: IncomingMessage): Promise<Json> {
  const chunks: Buffer[] = []
  for await (const chunk of req) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
  const raw = Buffer.concat(chunks).toString('utf8')
  if (!raw.trim()) return {}
  const parsed = JSON.parse(raw) as unknown
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw new Error('JSON object required')
  return parsed as Json
}

/** Local Vite/Node host for PDFium text, image, OCR, and fail-closed redaction proofs. */
export async function handlePdfNodeRequest(req: IncomingMessage, res: ServerResponse): Promise<boolean> {
  const url = req.url?.split('?')[0] ?? ''
  if (!url.startsWith(PDF_NODE_PREFIX)) return false
  if (req.method === 'OPTIONS') {
    res.statusCode = 204
    res.end()
    return true
  }
  if (req.method !== 'POST') {
    send(res, 405, { error: 'POST required' })
    return true
  }
  try {
    const body = await readJson(req)
    const bytes = bytesFrom(body.bytes)
    const page = typeof body.page === 'number' ? body.page : 1
    switch (url) {
      case `${PDF_NODE_PREFIX}/text-edit`: {
        const oldText = String(body.oldText ?? '')
        const newText = String(body.newText ?? '')
        const rect = numberTuple(body.rect, [0, 0, 612, 792])
        const result = await applyTextEdits(bytes, [{ page, oldText, newText, occurrence: 0, rect }])
        send(res, 200, { bytes: toB64(result.bytes), applied: result.applied, skipped: result.skipped })
        return true
      }
      case `${PDF_NODE_PREFIX}/text-insert`: {
        const text = String(body.text ?? '')
        const result = await applyTextInserts(bytes, [{ page, text, x: 72, y: 72, fontSize: 12 }])
        send(res, 200, { bytes: toB64(result.bytes), applied: result.applied, skipped: result.skipped })
        return true
      }
      case `${PDF_NODE_PREFIX}/images/list`: {
        const images = (await listPageImages(bytes)).filter((image) => image.page === page)
        send(res, 200, { images })
        return true
      }
      case `${PDF_NODE_PREFIX}/images/insert`: {
        const image = String(body.image ?? STAMP_PNG_B64)
        const rect = numberTuple(body.rect, [72, 72, 180, 144])
        const result = await applyImageEdits(bytes, [{ kind: 'insertImage', page, image, rect, layer: 'aboveText' }])
        send(res, 200, { bytes: toB64(result.bytes), skipped: result.skipped })
        return true
      }
      case `${PDF_NODE_PREFIX}/images/transform`: {
        const listed = (await listPageImages(bytes)).filter((image) => image.page === page)
        const first = listed[0]
        if (!first) throw new Error('no image on this page')
        const result = await applyImageEdits(bytes, [{
          kind: 'transformImage',
          page: first.page,
          oldRect: first.rect,
          rect: first.rect,
          quarterTurns: 1,
        }])
        send(res, 200, { bytes: toB64(result.bytes), skipped: result.skipped })
        return true
      }
      case `${PDF_NODE_PREFIX}/images/delete`: {
        const listed = (await listPageImages(bytes)).filter((image) => image.page === page)
        const first = listed[0]
        if (!first) throw new Error('no image on this page')
        const result = await applyImageEdits(bytes, [{ kind: 'deleteImage', page: first.page, oldRect: first.rect }])
        send(res, 200, { bytes: toB64(result.bytes), skipped: result.skipped })
        return true
      }
      case `${PDF_NODE_PREFIX}/ocr/render`: {
        const rendered = await renderPageToPng(bytes, page, 1)
        send(res, 200, { png: Buffer.from(rendered.png).toString('base64'), width: rendered.width, height: rendered.height })
        return true
      }
      case `${PDF_NODE_PREFIX}/ocr/layer`: {
        const result = await applyOcrLayer(bytes, async () => [{ text: String(body.text ?? 'OCR'), box: [0.1, 0.1, 0.4, 0.14] }])
        send(res, 200, {
          bytes: toB64(result.bytes),
          pagesOcred: result.pagesOcred,
          skipped: result.pagesSkippedExistingText,
          failed: result.failed,
        })
        return true
      }
      case `${PDF_NODE_PREFIX}/redact/text`: {
        const rect = numberTuple(body.rect, [72, 650, 360, 730])
        const result = await applyWholeTextRedactions(bytes, [{ page, rect }])
        send(res, 200, { bytes: toB64(result.bytes), proofs: result.proofs })
        return true
      }
      case `${PDF_NODE_PREFIX}/redact/image`: {
        const listed = (await listPageImages(bytes)).filter((image) => image.page === page)
        const first = listed[0]
        if (!first) throw new Error('no image on this page')
        const result = await applyWholeImageRedactions(bytes, [{ page: first.page, rect: first.rect }])
        send(res, 200, { bytes: toB64(result.bytes), proofs: result.proofs })
        return true
      }
      case `${PDF_NODE_PREFIX}/redact/annotation`: {
        const targets = (await listWholeAnnotationRedactionTargets(bytes)).filter((target) => target.page === page)
        const first = targets[0]
        if (!first) throw new Error('no redactable annotation on this page')
        const next = await applyWholeAnnotationRedactions(bytes, [first])
        send(res, 200, { bytes: toB64(next) })
        return true
      }
      case `${PDF_NODE_PREFIX}/redact/path`: {
        const paths = (await listPagePaths(bytes)).filter((item) => item.page === page)
        const first = paths[0]
        if (!first) throw new Error('no path object on this page')
        const result = await applyWholePathRedactions(bytes, [{ page: first.page, rect: first.rect }])
        send(res, 200, { bytes: toB64(result.bytes), proofs: result.proofs })
        return true
      }
      case `${PDF_NODE_PREFIX}/redact/xobject`: {
        const objects = (await listPageXObjects(bytes)).filter((item) => item.page === page)
        const first = objects[0]
        if (!first) throw new Error('no form XObject on this page')
        const result = await applyWholeXObjectRedactions(bytes, [{
          page: first.page,
          name: first.name,
          rect: first.rect,
          wholeXObject: true,
        }])
        send(res, 200, { bytes: toB64(result.bytes), proofs: result.proofs })
        return true
      }
      case `${PDF_NODE_PREFIX}/redact/form`: {
        const fields = await listRedactableFormFields(bytes)
        const first = fields[0]
        if (!first) throw new Error('no redactable form field')
        const result = await redactFormFields(bytes, [first])
        send(res, 200, { bytes: toB64(result.bytes), removed: result.removed })
        return true
      }
      default:
        send(res, 404, { error: `unknown PDF host route ${url}` })
        return true
    }
  } catch (error) {
    send(res, 400, { error: error instanceof Error ? error.message : String(error) })
    return true
  }
}
