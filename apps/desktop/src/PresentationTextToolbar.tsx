import type { ReactNode } from 'react';
import type { PptxNativeExactParagraphV1, PptxNativeExactTextRunV1 } from '@injoffice/pptx-wasm';
import FormattingToolbar, { type FormattingPatch, type FormattingSection, type FormattingValues } from './FormattingToolbar';
import { RibbonButton } from './Ribbon';

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

/**
 * Slide text formatting over the shared FormattingToolbar. `section="font"` and
 * `section="paragraph"` render one ribbon group each (Home › Font, Home › Paragraph);
 * the default renders the whole toolbar. Unsupported Office controls stay disabled
 * with their reason: Underline sits with the font controls, Bullets with paragraph.
 */
export default function PresentationTextToolbar({ run, align, disabled, onRunChange, onAlignChange, children, section = 'all' }: {
  /** The selected text segment (a draft run when one is pending). Undefined when no editable text is selected. */
  run?: PptxNativeExactTextRunV1; align?: PresentationAlignment; disabled: boolean;
  onRunChange(patch: PresentationRunPatch): void; onAlignChange(align: PresentationAlignment): void; children?: ReactNode; section?: FormattingSection;
}) {
  const values = presentationTextValues(run, align);
  const underline = <RibbonButton icon="underline" label="Underline" labelHidden title={PRESENTATION_TEXT_UNSUPPORTED.underline} disabled aria-pressed={false} />;
  const bullets = <RibbonButton icon="list" label="Bullets" aria-label="Bullets" title={PRESENTATION_TEXT_UNSUPPORTED.bullets} disabled aria-pressed={false} />;
  return <div className={`presentation-text-toolbar presentation-text-toolbar-${section}`} aria-label={section === 'font' ? 'Slide font' : section === 'paragraph' ? 'Slide paragraph' : 'Slide text formatting'}>
    <FormattingToolbar kind="pptx" section={section} disabled={disabled} values={values} onChange={patch => { const next = presentationTextPatch(patch); if (next.run) onRunChange(next.run); if (next.align) onAlignChange(next.align); }}>
      {section === 'all' && <div className="formatting-group">{underline}{bullets}</div>}
      {section === 'font' && underline}
      {section === 'paragraph' && bullets}
      {children}
    </FormattingToolbar>
  </div>;
}
