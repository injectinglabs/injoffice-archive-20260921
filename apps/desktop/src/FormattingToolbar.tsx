import type { ReactNode } from 'react';
import { ColorButton, RibbonButton, RibbonCombo, RibbonRows, type RibbonColor } from './Ribbon'

export interface FormattingValues {
  font?: string; size?: number; bold?: boolean; italic?: boolean; underline?: boolean;
  color?: string; alignment?: string; characterEditable?: boolean; fill?: string; numberFormat?: string;
}
export type FormattingPatch = Partial<FormattingValues>;
/** `font` and `paragraph` render one Office ribbon group each; `all` is the classic single toolbar. */
export type FormattingSection = 'all' | 'font' | 'paragraph';

/** Office's font-size ladder, shared by the ribbon combo and the mini toolbar's grow/shrink buttons. */
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72];

/** The next size up (`1`) or down (`-1`) the ladder, or undefined when the size in effect is unknown or at the end. */
export function stepFontSize(size: number | undefined, direction: 1 | -1): number | undefined {
  if (size === undefined || !Number.isFinite(size)) return undefined;
  const ladder = [...new Set([size, ...FONT_SIZES])].sort((a, b) => a - b);
  return ladder[ladder.indexOf(size) + direction];
}

/** The text-colour palette shared by the ribbon, the mini toolbar and the slide inspector. */
export const TEXT_COLORS: RibbonColor[] = [['#000000', 'Black'], ['#20242B', 'Dark gray'], ['#687386', 'Gray'], ['#2459AD', 'Blue'], ['#B33C3C', 'Red'], ['#217447', 'Green']];

/** The alignment buttons Office shows for each host, in ribbon order. */
export function alignmentOptions(kind: 'docx' | 'xlsx' | 'pptx'): Array<{ value: string; label: string; icon: 'alignLeft' | 'alignCenter' | 'alignRight' | 'alignJustify' | 'alignDistribute' }> {
  const base = [
    { value: 'left', label: 'Align left', icon: 'alignLeft' as const },
    { value: 'center', label: 'Center', icon: 'alignCenter' as const },
    { value: 'right', label: 'Align right', icon: 'alignRight' as const },
  ];
  return kind === 'docx' ? [...base, { value: 'both', label: 'Justify', icon: 'alignJustify' as const }, { value: 'distribute', label: 'Distribute', icon: 'alignDistribute' as const }] : base;
}

/** Status text describing what the formatting controls currently apply to. */
export function formattingScope(kind: 'docx' | 'xlsx' | 'pptx', values: FormattingValues | undefined, scopeLabel?: string): string {
  if (!values) return kind === 'xlsx' ? 'Select a cell to format' : 'Select text to format';
  return kind === 'xlsx' ? 'Selected cell' : kind === 'docx' ? scopeLabel ?? 'Selected segment · direct formatting' : 'Selected text segment';
}

/**
 * The Font and Paragraph ribbon groups. Every control shows the value in effect
 * at the caret; where it is unknown or mixed across the selection the control is
 * empty (Word's behaviour) rather than carrying a placeholder word.
 */
export default function FormattingToolbar({ kind, values, disabled, onChange, scopeLabel, section = 'all', children }: {
  scopeLabel?:string; kind: 'docx' | 'xlsx' | 'pptx'; values?: FormattingValues; disabled: boolean; onChange(patch: FormattingPatch): void; section?: FormattingSection; children?: ReactNode;
}) {
  const inactive = disabled || !values;
  const characterInactive = inactive || values?.characterEditable === false;
  const fonts = [...new Set([values?.font || 'Arial', 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Times New Roman', 'Verdana', 'DejaVu Sans'])];
  const sizes = [...new Set([values?.size || 11, ...FONT_SIZES])].sort((a, b) => a - b);
  const family = <RibbonCombo className="ribbon-combo-font" label="Font family" disabled={characterInactive} value={values?.font ?? ''} options={fonts.map(font => ({ value: font, label: font }))} onChange={font => onChange({ font })} />;
  const size = <RibbonCombo className="ribbon-combo-size" label="Font size" disabled={characterInactive} value={values?.size === undefined ? '' : String(values.size)} options={sizes.map(size => ({ value: String(size), label: String(size) }))} onChange={size => onChange({ size: Number(size) })} />;
  // Only the DOCX editor binds ⌘B/⌘I/⌘U; other hosts get no shortcut hint they cannot honour.
  const keyed = kind === 'docx';
  const emphasis = <>
    <RibbonButton icon="bold" label="Bold" shortcut={keyed ? 'bold' : undefined} labelHidden aria-pressed={values?.bold ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ bold: !values?.bold })} />
    <RibbonButton icon="italic" label="Italic" shortcut={keyed ? 'italic' : undefined} labelHidden aria-pressed={values?.italic ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ italic: !values?.italic })} />
    {kind === 'docx' && <RibbonButton icon="underline" label="Underline" shortcut="underline" labelHidden aria-pressed={values?.underline ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ underline: !values?.underline })} />}
  </>;
  const color = <ColorButton label="Text color" icon="fontColor" disabled={characterInactive} value={values?.color?.toUpperCase()} colors={TEXT_COLORS} onChange={color => onChange({ color })} />;
  const alignment = <span className="formatting-alignment">
    {alignmentOptions(kind).map(option => <RibbonButton key={option.value} icon={option.icon} label={option.label} labelHidden aria-pressed={values?.alignment === option.value} disabled={inactive} onClick={() => onChange({ alignment: option.value })} />)}
  </span>;
  if (section === 'font') return <div className="formatting-toolbar formatting-toolbar-font" aria-label="Font formatting"><RibbonRows><div className="formatting-group">{family}{size}</div><div className="formatting-group">{emphasis}{children}{color}</div></RibbonRows></div>;
  if (section === 'paragraph') return <div className="formatting-toolbar formatting-toolbar-paragraph" aria-label="Paragraph alignment"><div className="formatting-group">{alignment}</div>{children}</div>;
  return <div className="formatting-toolbar" aria-label="Formatting">
    <div className="formatting-group">{family}{size}{emphasis}{color}</div>
    <div className="formatting-group">{alignment}</div>
    {children}
    <span className="formatting-scope">{formattingScope(kind, values, scopeLabel)}</span>
  </div>;
}
