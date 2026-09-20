import type { ReactNode } from 'react';
import type { PptxNativeExactParagraphV1, PptxNativeExactTextRunV1 } from '@injoffice/pptx-wasm';
import FormattingToolbar, { type FormattingPatch, type FormattingValues } from './FormattingToolbar';

/** Run fields the native PPTX `text.replace` transaction accepts (packages/pptx-wasm validateParagraph). */
export type PresentationRunPatch = Partial<Pick<PptxNativeExactTextRunV1, 'fontFamily' | 'fontSizeHundredthPt' | 'bold' | 'italic' | 'color'>>;
export type PresentationAlignment = PptxNativeExactParagraphV1['align'];
/** Controls Office has that the native transaction refuses; shown disabled with the reason, never wired to a no-op. */
export const PRESENTATION_TEXT_UNSUPPORTED = {
  underline: 'Underline is not supported by the native PPTX transaction',
  bullets: 'Bullets are not supported by the native PPTX transaction',
} as const;

export function presentationTextValues(run: PptxNativeExactTextRunV1 | undefined, align: PresentationAlignment | undefined): FormattingValues | undefined {
  if (!run) return undefined;
  return { font: run.fontFamily, size: Number.isFinite(run.fontSizeHundredthPt) ? run.fontSizeHundredthPt / 100 : undefined, bold: run.bold, italic: run.italic, color: `#${run.color}`, alignment: align };
}
/** Translate the shared toolbar patch into exact run fields; alignment is a paragraph property and is returned separately. */
export function presentationTextPatch(patch: FormattingPatch): { run?: PresentationRunPatch; align?: PresentationAlignment } {
  const run: PresentationRunPatch = {};
  if (patch.font !== undefined) run.fontFamily = patch.font;
  if (patch.size !== undefined) { const size = Math.round(patch.size * 100); if (Number.isSafeInteger(size) && size >= 1 && size <= 400_000) run.fontSizeHundredthPt = size; }
  if (patch.bold !== undefined) run.bold = patch.bold;
  if (patch.italic !== undefined) run.italic = patch.italic;
  if (patch.color !== undefined) { const color = patch.color.replace(/^#/, '').toUpperCase(); if (/^[0-9A-F]{6}$/.test(color)) run.color = color; }
  const align = patch.alignment === 'left' || patch.alignment === 'center' || patch.alignment === 'right' ? patch.alignment : undefined;
  return { ...(Object.keys(run).length ? { run } : {}), ...(align ? { align } : {}) };
}

export default function PresentationTextToolbar({ run, align, disabled, onRunChange, onAlignChange, children }: {
  /** The selected text segment (a draft run when one is pending). Undefined when no editable text is selected. */
  run?: PptxNativeExactTextRunV1; align?: PresentationAlignment; disabled: boolean;
  onRunChange(patch: PresentationRunPatch): void; onAlignChange(align: PresentationAlignment): void; children?: ReactNode;
}) {
  const values = presentationTextValues(run, align);
  return <div className="presentation-text-toolbar" aria-label="Slide text formatting">
    <FormattingToolbar kind="pptx" disabled={disabled} values={values} onChange={patch => { const next = presentationTextPatch(patch); if (next.run) onRunChange(next.run); if (next.align) onAlignChange(next.align); }}>
      <div className="formatting-group">
        <button aria-label="Underline" title={PRESENTATION_TEXT_UNSUPPORTED.underline} disabled aria-pressed={false}><u>U</u></button>
        <button aria-label="Bullets" title={PRESENTATION_TEXT_UNSUPPORTED.bullets} disabled aria-pressed={false}>Bullets</button>
      </div>
      {children}
    </FormattingToolbar>
  </div>;
}
