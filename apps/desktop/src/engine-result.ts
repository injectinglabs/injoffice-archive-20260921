/**
 * Engine refusals that are not failures.
 *
 * Re-applying a value the document already carries (picking the colour that is already applied,
 * setting a style the paragraph already has) makes the native engines refuse with `SEMANTIC_NO_OP`:
 * the transaction produces no change. Nothing went wrong, so the renderer must not raise an error
 * banner for it — the UI already shows the value the user asked for.
 */
export const SEMANTIC_NO_OP = 'SEMANTIC_NO_OP'

/** Issue codes carried by a rejected engine result, including its `issues` list and `cause` chain. */
function issueCodes(reason: unknown, depth = 0): string[] {
  if (depth > 4 || !reason || typeof reason !== 'object') return []
  const record = reason as { code?: unknown; issues?: unknown; cause?: unknown }
  const codes: string[] = []
  if (typeof record.code === 'string') codes.push(record.code)
  if (Array.isArray(record.issues)) {
    for (const issue of record.issues) {
      if (issue && typeof issue === 'object' && typeof (issue as { code?: unknown }).code === 'string') codes.push((issue as { code: string }).code)
    }
  }
  if (record.cause) codes.push(...issueCodes(record.cause, depth + 1))
  return codes
}

/** True when the engine refused only because the request changes nothing. */
export function isSemanticNoOp(reason: unknown): boolean {
  const codes = issueCodes(reason)
  if (codes.length) return codes.every(code => code === SEMANTIC_NO_OP)
  const message = reason instanceof Error ? reason.message : typeof reason === 'string' ? reason : ''
  return new RegExp(`(^|[^A-Z_])${SEMANTIC_NO_OP}([^A-Z_]|$)`).test(message)
}

/** Message to show for a rejected engine call; '' for a no-op, which is success with nothing to do. */
export function engineErrorMessage(reason: unknown): string {
  if (isSemanticNoOp(reason)) return ''
  return reason instanceof Error ? reason.message : String(reason)
}
