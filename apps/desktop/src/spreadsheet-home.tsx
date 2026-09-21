// Excel's Home controls for the XLSX ribbon: alignment toggle icons, colour
// buttons with the colour underline, a Borders icon menu and the Cells group.
// Styling lives in spreadsheet.css (imported by SpreadsheetEditor) so this file
// stays CSS-free for the unit-test bundlers.
import type { ReactNode } from 'react';
import type { StyleDelta } from '@injoffice/sheets/browser';

type HorizontalAlignment = NonNullable<StyleDelta['horizontal_alignment']>;
type VerticalAlignment = NonNullable<StyleDelta['vertical_alignment']>;

const box = 'M4 4h16v16H4z';
/** 24×24 stroke glyphs for the controls Excel draws as icons. */
const glyphs = {
  alignLeft: ['M4 6h16M4 10h9M4 14h16M4 18h9'],
  alignCenter: ['M4 6h16M7.5 10h9M4 14h16M7.5 18h9'],
  alignRight: ['M4 6h16M11 10h9M4 14h16M11 18h9'],
  alignTop: [box, 'M7 8h10M7 11.5h6'],
  alignMiddle: [box, 'M7 10.5h10M7 14h6'],
  alignBottom: [box, 'M7 13h10M7 16.5h6'],
  textColor: ['M6.5 15.5 11.5 5l5 10.5M8.5 12.5h6'],
  fillColor: ['M11 4 5.5 9.5 12 16l6.5-6.5z', 'M20.5 14.5c.9 1.3.9 3-.7 3s-1.6-1.7-.7-3z'],
  borders: [box, 'M4 12h16M12 4v16'],
  bordersOuter: [box],
  bordersBottom: ['M4 20h16', 'M4 4h16v16H4z'],
  bordersNone: ['M4 4h16v16H4z'],
  cellsInsert: ['M4 4h11v11H4z', 'M18 13v8M14 17h8'],
  cellsDelete: ['M4 4h11v11H4z', 'M14 17h8'],
  cellsFormat: ['M4 4h11v11H4z', 'M16 21l5-5M15 22l1.2-3.2 3.2-1.2'],
} as const;
export type SheetIconName = keyof typeof glyphs;

export function SheetIcon({ name }: { name: SheetIconName }) {
  return <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
    {glyphs[name].map(path => <path key={path} d={path} />)}
  </svg>;
}

interface ToggleProps { icon: SheetIconName; label: string; title?: string; pressed: boolean; disabled?: boolean; onClick(): void }
/** Icon-only ribbon command that reports its pressed state, like Excel's toggles. */
export function SheetToggle({ icon, label, title, pressed, disabled, onClick }: ToggleProps) {
  return <button type="button" className="ribbon-button ribbon-button-icon-only" title={title ?? label} aria-label={label} aria-pressed={pressed} disabled={disabled} onClick={onClick}>
    <SheetIcon name={icon} />
  </button>;
}

export const HORIZONTAL_ALIGNMENTS: { value: HorizontalAlignment; icon: SheetIconName; label: string }[] = [
  { value: 'left', icon: 'alignLeft', label: 'Align left' },
  { value: 'center', icon: 'alignCenter', label: 'Center' },
  { value: 'right', icon: 'alignRight', label: 'Align right' },
];
export const VERTICAL_ALIGNMENTS: { value: VerticalAlignment; icon: SheetIconName; label: string }[] = [
  { value: 'top', icon: 'alignTop', label: 'Top align' },
  { value: 'middle', icon: 'alignMiddle', label: 'Middle align' },
  { value: 'bottom', icon: 'alignBottom', label: 'Bottom align' },
];

/** Colour command: the glyph over an underline in the current colour, opening the OS picker. */
export function SheetColorButton({ icon, label, value, disabled, onChange }: {
  icon: SheetIconName; label: string; value: string; disabled?: boolean; onChange(color: string): void;
}) {
  return <span className={`sheet-color-button${disabled ? ' is-disabled' : ''}`} title={label}>
    <SheetIcon name={icon} />
    <span className="sheet-color-underline" style={{ background: value }} aria-hidden="true" />
    <input type="color" aria-label={label} disabled={disabled} value={value} onChange={event => onChange(event.target.value)} />
  </span>;
}

export type SheetBorderPreset = 'all' | 'outer' | 'bottom' | 'clear';
export const BORDER_PRESETS: { value: SheetBorderPreset; icon: SheetIconName; label: string }[] = [
  { value: 'all', icon: 'borders', label: 'All borders' },
  { value: 'outer', icon: 'bordersOuter', label: 'Outside borders' },
  { value: 'bottom', icon: 'bordersBottom', label: 'Bottom border' },
  { value: 'clear', icon: 'bordersNone', label: 'No border' },
];

/** Borders as an icon menu. `reason` disables it and explains why in the tooltip. */
export function SheetBordersMenu({ disabled, reason, onApply, line = 'thin', color = '#000000', onLine, onColor }: { disabled?: boolean; reason?: string; onApply(preset: SheetBorderPreset): void; line?: NonNullable<StyleDelta['border_top']>['style']; color?: string; onLine?(line: NonNullable<StyleDelta['border_top']>['style']): void; onColor?(color: string): void }) {
  const title = reason ?? 'Borders';
  if (disabled || reason) return <button type="button" className="ribbon-button ribbon-button-icon-only" title={title} aria-label="Borders" disabled><SheetIcon name="borders" /></button>;
  return <details className="sheet-menu sheet-borders-menu">
    <summary title={title} aria-label="Borders"><SheetIcon name="borders" /><span className="sheet-menu-caret" aria-hidden="true">▾</span></summary>
    <div><label>Line style<select aria-label="Border line style" value={line} onChange={event => onLine?.(event.target.value as NonNullable<StyleDelta['border_top']>['style'])}>{(['thin','medium','thick','double','dotted','dashed'] as const).map(value => <option key={value}>{value}</option>)}</select></label><label>Line color<input aria-label="Border color" type="color" value={color} onChange={event => onColor?.(event.target.value)} /></label>{BORDER_PRESETS.map(preset => <button key={preset.value} type="button" title={preset.label} onClick={() => onApply(preset.value)}><SheetIcon name={preset.icon} />{preset.label}</button>)}</div>
  </details>;
}

/** An icon command with a caret that opens a small menu, like Excel's Insert/Delete/Format. */
export function SheetMenuButton({ icon, label, disabled, reason, children }: { icon: SheetIconName; label: string; disabled?: boolean; reason?: string; children: ReactNode }) {
  if (disabled || reason) return <button type="button" className="ribbon-button ribbon-button-icon-only" title={reason ?? label} aria-label={label} disabled><SheetIcon name={icon} /></button>;
  return <details className="sheet-menu">
    <summary title={label} aria-label={label}><SheetIcon name={icon} /><span className="sheet-menu-caret" aria-hidden="true">▾</span></summary>
    <div>{children}</div>
  </details>;
}
