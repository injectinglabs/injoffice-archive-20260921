import type { AnnotDeleteSpec, FormValueSpec, MarkupSpec, NoteEditSpec } from './annotate/index.js'

/** The only PDF mutations safe for the initial opaque collaboration log. */
export type PdfCollabOperation =
  | { kind: 'annotation.markup'; value: MarkupSpec }
  | { kind: 'annotation.note'; value: NoteEditSpec }
  | { kind: 'annotation.delete'; value: AnnotDeleteSpec }
  | { kind: 'form.value'; value: FormValueSpec }

export type PdfCollabDecode = { ok: true; operation: PdfCollabOperation } | { ok: false; reason: string }

/** Decode untrusted opaque collab-log values. Page/content/image/structural
 * mutations deliberately have no representation and are rejected. */
export function decodePdfCollabOperation(value: unknown): PdfCollabDecode {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return { ok: false, reason: 'operation must be an object' }
  const op = value as { kind?: unknown; value?: unknown }
  if (typeof op.kind !== 'string' || !('value' in op)) return { ok: false, reason: 'operation kind and value are required' }
  switch (op.kind) {
    case 'annotation.markup': case 'annotation.note': case 'annotation.delete': case 'form.value':
      if (!op.value || typeof op.value !== 'object' || Array.isArray(op.value)) return { ok: false, reason: 'operation value must be an object' }
      return { ok: true, operation: { kind: op.kind, value: op.value } as PdfCollabOperation }
    default: return { ok: false, reason: `unsupported PDF collaboration operation: ${op.kind}` }
  }
}

export function encodePdfCollabOperation(operation: PdfCollabOperation): PdfCollabOperation {
  const decoded = decodePdfCollabOperation(operation)
  if (!decoded.ok) throw new Error(decoded.reason)
  return decoded.operation
}
