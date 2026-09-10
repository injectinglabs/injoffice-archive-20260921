import { createDocxWasmClient } from '@injoffice/docx-wasm'
import { adaptNativeDocxProseMirrorTransactionV1 } from '@injoffice/docs/native-docx'

type Transaction = Parameters<typeof adaptNativeDocxProseMirrorTransactionV1>[1]
type NativeDocument = Awaited<ReturnType<ReturnType<typeof createDocxWasmClient>['extract']>>

/** The host projects its editor transaction against this exact source. */
export async function applyDocumentTransaction(
  original: Uint8Array,
  project: (document: NativeDocument) => Transaction,
) {
  const client = createDocxWasmClient()
  try {
    const document = await client.extract(original)
    const adapted = adaptNativeDocxProseMirrorTransactionV1(document, project(document))
    if (!adapted.ok) throw new Error(JSON.stringify(adapted.issues))
    const output = await client.apply(original, document, adapted.value.envelope)
    const readback = await client.extract(output)
    return { output, readback }
  } finally {
    client.terminate()
  }
}
