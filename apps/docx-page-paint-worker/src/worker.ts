import { stdin, stdout } from 'node:process'
import {
  DOCX_PAGE_PAINT_WORKER_MAX_FRAME_BYTES,
  dispatchNativeDocxPagePaintWorkerRequestV1,
  encodeNativeDocxPagePaintWorkerResponseV1,
} from './protocol.js'

let pending = Buffer.alloc(0)
let chain = Promise.resolve()
let ended = false
const args = process.argv.slice(2)
if (args.length !== 0 && (args.length !== 2 || args[0] !== '--font-manifest')) fatal('invalid operator font arguments')
const hostFontManifestPath = args[1]

function fatal(message: string): never {
  process.stderr.write(`${message.replace(/[\u0000\r\n]/g, ' ').slice(0, 1_024)}\n`)
  process.exit(1)
}

function parse(): void {
  while (pending.byteLength >= 4) {
    const length = pending.readUInt32BE(0)
    if (length === 0 || length > DOCX_PAGE_PAINT_WORKER_MAX_FRAME_BYTES) fatal('native page-paint worker received an invalid frame length')
    if (pending.byteLength < 4 + length) return
    const payload = Buffer.from(pending.subarray(4, 4 + length))
    pending = Buffer.from(pending.subarray(4 + length))
    chain = chain.then(async () => {
      let request: unknown
      try { request = JSON.parse(payload.toString('utf8')) } catch { request = null }
      const response = await dispatchNativeDocxPagePaintWorkerRequestV1(request, hostFontManifestPath)
      const frame = encodeNativeDocxPagePaintWorkerResponseV1(response)
      await new Promise<void>((resolve, reject) => stdout.write(frame, (error) => error ? reject(error) : resolve()))
    }).catch((error: unknown) => fatal(error instanceof Error ? error.message : 'native page-paint worker failed'))
  }
  if (ended && pending.byteLength !== 0) fatal('native page-paint worker received a truncated final frame')
}

stdin.on('data', (chunk: Buffer) => {
  if (pending.byteLength + chunk.byteLength > DOCX_PAGE_PAINT_WORKER_MAX_FRAME_BYTES + 4) fatal('native page-paint worker input buffer exceeded its bound')
  pending = Buffer.concat([pending, chunk])
  parse()
})
stdin.on('end', () => { ended = true; parse(); void chain.then(() => process.exit(0)) })
stdin.on('error', (error) => fatal(error.message))
