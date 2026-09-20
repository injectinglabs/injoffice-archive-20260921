import type { ReactNode } from 'react';
import { RibbonButton, RibbonRows } from './Ribbon'

export interface FormattingValues {
  font?: string; size?: number; bold?: boolean; italic?: boolean; underline?: boolean;
  color?: string; alignment?: string; characterEditable?: boolean; fill?: string; numberFormat?: string;
}
export type FormattingPatch = Partial<FormattingValues>;
/** `font` and `paragraph` render one Office ribbon group each; `all` is the classic single toolbar. */
export type FormattingSection = 'all' | 'font' | 'paragraph';

/** Status text describing what the formatting controls currently apply to. */
export function formattingScope(kind: 'docx' | 'xlsx' | 'pptx', values: FormattingValues | undefined, scopeLabel?: string): string {
  if (!values) return kind === 'xlsx' ? 'Select a cell to format' : 'Select text to format';
  return kind === 'xlsx' ? 'Selected cell' : kind === 'docx' ? scopeLabel ?? 'Selected segment · direct formatting' : 'Selected text segment';
}

export default function FormattingToolbar({ kind, values, disabled, onChange, scopeLabel, section = 'all', children }: {
  scopeLabel?:string; kind: 'docx' | 'xlsx' | 'pptx'; values?: FormattingValues; disabled: boolean; onChange(patch: FormattingPatch): void; section?: FormattingSection; children?: ReactNode;
}) {
  const inactive = disabled || !values;
  const characterInactive = inactive || values?.characterEditable === false;
  const fonts = [...new Set([values?.font || 'Arial', 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Times New Roman', 'Verdana', 'DejaVu Sans'])];
  const sizes = [...new Set([values?.size || 11, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72])].sort((a, b) => a - b);
  const colors = [['#20242B', 'Dark gray'], ['#000000', 'Black'], ['#2459AD', 'Blue'], ['#B33C3C', 'Red'], ['#217447', 'Green'], ['#687386', 'Gray']];
  if (values?.color && !colors.some(([color]) => color === values.color?.toUpperCase())) colors.unshift([values.color.toUpperCase(), 'Current color']);
  const family = <select aria-label="Font family" title="Font family" disabled={characterInactive} value={values?.font || ''} onChange={event => onChange({ font: event.target.value })}>{!values?.font && <option value="">{scopeLabel?'Mixed / inherited font':'Inherited font'}</option>}{fonts.map(font => <option key={font}>{font}</option>)}</select>;
  const size = <select aria-label="Font size" title="Font size" disabled={characterInactive} value={values?.size || ''} onChange={event => onChange({ size: Number(event.target.value) })}>{!values?.size && <option value="">{scopeLabel?'Mixed / inherited size':'Inherited size'}</option>}{sizes.map(size => <option key={size} value={size}>{size}</option>)}</select>;
  const emphasis = <>
    <RibbonButton icon="bold" label="Bold" shortcut="bold" labelHidden aria-pressed={values?.bold ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ bold: !values?.bold })} />
    <RibbonButton icon="italic" label="Italic" shortcut="italic" labelHidden aria-pressed={values?.italic ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ italic: !values?.italic })} />
    {kind === 'docx' && <RibbonButton icon="underline" label="Underline" shortcut="underline" labelHidden aria-pressed={values?.underline ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ underline: !values?.underline })} />}
  </>;
  const color = <select aria-label="Text color" title="Text color" disabled={characterInactive} value={values?.color?.toUpperCase() || ''} onChange={event => onChange({ color: event.target.value })}>{!values?.color && <option value="">{scopeLabel?'Mixed / inherited color':'Inherited color'}</option>}{colors.map(([color, name]) => <option key={color} value={color}>{name}</option>)}</select>;
  const alignment = <select aria-label="Text alignment" title="Text alignment" disabled={inactive} value={values?.alignment || ''} onChange={event => onChange({ alignment: event.target.value })}>
    {!values?.alignment && <option value="">Inherited alignment</option>}{kind === 'xlsx' && <option value="general">General alignment</option>}<option value="left">Align left</option><option value="center">Center</option><option value="right">Align right</option>{kind === 'docx' && <><option value="both">Justify</option><option value="distribute">Distribute</option></>}
  </select>;
  if (section === 'font') return <div className="formatting-toolbar formatting-toolbar-font" aria-label="Font formatting"><RibbonRows><div className="formatting-group">{family}{size}</div><div className="formatting-group">{emphasis}{color}</div></RibbonRows>{children}</div>;
  if (section === 'paragraph') return <div className="formatting-toolbar formatting-toolbar-paragraph" aria-label="Paragraph alignment"><div className="formatting-group">{alignment}</div>{children}</div>;
  return <div className="formatting-toolbar" aria-label="Formatting">
    <div className="formatting-group">{family}{size}{emphasis}{color}</div>
    <div className="formatting-group">{alignment}</div>
    {children}
    <span className="formatting-scope">{formattingScope(kind, values, scopeLabel)}</span>
  </div>;
}
