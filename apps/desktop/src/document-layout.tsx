import { useEffect, useRef, useState, type ReactNode } from 'react'
import type { NativeDocxSectionV1 } from '../../../packages/docs/src/nativeContract'
import { RibbonButton } from './Ribbon'
import type { RibbonIconName } from './RibbonIcons'

/**
 * Page setup and paragraph patches as the engine takes them. Declared here rather than imported
 * from the older controls: those modules are outside `npm run typecheck` (tsconfig exclusions),
 * and importing them would drag their errors into this file's program.
 */
export type PagePatch = { width_twips: number; height_twips: number; orientation: 'portrait' | 'landscape'; margin_top_twips: number; margin_right_twips: number; margin_bottom_twips: number; margin_left_twips: number }
export type LayoutParagraphPatch = { indent_left_twips?: number | null; indent_right_twips?: number | null; spacing_before_twips?: number | null; spacing_after_twips?: number | null; first_line_twips?: number | null; hanging_twips?: number | null; line_spacing?: number | null; line_rule?: 'auto' | 'exact' | 'atLeast' | null; outline_level?: number | null }
export interface LayoutParagraphSettings { indent_left_twips?: number; indent_right_twips?: number; spacing_before_twips?: number; spacing_after_twips?: number; first_line_twips?: number; hanging_twips?: number; line_spacing?: number; line_rule?: 'auto' | 'exact' | 'atLeast'; outline_level?: number }
/** The section as the engine reports it, with the edit policy the native contract does not model yet. */
type LayoutSection = NativeDocxSectionV1 & { edit_policy?: { allowed_operations: string[] } }
function pagePatch(section: LayoutSection): PagePatch {
  const page = section.page
  return { width_twips: page.width_twips, height_twips: page.height_twips, orientation: page.orientation, margin_top_twips: page.margins.top_twips, margin_right_twips: page.margins.right_twips, margin_bottom_twips: page.margins.bottom_twips, margin_left_twips: page.margins.left_twips }
}

/** Page setup reaches the engine as one section patch, so it is a whole-document change here. */
const SCOPE_NOTE = 'Applies to the whole document in this build'
const inches = (twips: number) => `${Number((twips / 1440).toFixed(2))}"`
const papers = [
  { name: 'Letter', width: 12240, height: 15840 },
  { name: 'A4', width: 11906, height: 16838 },
  { name: 'Legal', width: 12240, height: 20160 },
  { name: 'A5', width: 8391, height: 11906 },
]
const marginPresets = [
  { name: 'Normal', top: 1440, bottom: 1440, left: 1440, right: 1440 },
  { name: 'Narrow', top: 720, bottom: 720, left: 720, right: 720 },
  { name: 'Moderate', top: 1440, bottom: 1440, left: 1080, right: 1080 },
  { name: 'Wide', top: 1440, bottom: 1440, left: 2880, right: 2880 },
]

interface MenuItem { id: string; label: string; detail?: string; selected: boolean; apply(): void }

/** A ribbon gallery button: the current value on the face, the choices in a small menu. */
function LayoutMenu({ icon, label, value, disabled, items }: { icon: RibbonIconName; label: string; value: string; disabled: boolean; items: MenuItem[] }) {
  const [open, setOpen] = useState(false)
  const box = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const outside = (event: Event) => { if (!box.current?.contains(event.target as Node)) setOpen(false) }
    const escape = (event: KeyboardEvent) => { if (event.key === 'Escape') setOpen(false) }
    window.document.addEventListener('mousedown', outside)
    window.document.addEventListener('keydown', escape)
    return () => { window.document.removeEventListener('mousedown', outside); window.document.removeEventListener('keydown', escape) }
  }, [open])
  const title = disabled ? `${label}: ${value}. ${SCOPE_NOTE}; this document's section structure is read only here.` : `${label}: ${value}. ${SCOPE_NOTE}.`
  return <div className="document-layout-menu" ref={box}>
    <RibbonButton icon={icon} label={label} title={title} disabled={disabled} aria-expanded={open} aria-haspopup="menu" onClick={() => setOpen(value => !value)}>
      <span className="document-layout-value">{value}</span>
      <span className="document-layout-caret" aria-hidden="true">▾</span>
    </RibbonButton>
    {open && <div className="document-layout-list" role="menu" aria-label={label}>
      {items.map(item => <button key={item.id} type="button" role="menuitemradio" aria-checked={item.selected} onClick={() => { item.apply(); setOpen(false) }}>
        <span>{item.label}</span>{item.detail && <small>{item.detail}</small>}
      </button>)}
    </div>}
  </div>
}

/** Word's Layout → Page Setup: Margins, Orientation and Size, each showing the current value. */
export function DocumentPageSetup({ section, disabled, onChange }: { section?: LayoutSection; disabled: boolean; onChange(patch: PagePatch): void }): ReactNode {
  if (!section) return null
  const readOnly = disabled || !section.edit_policy?.allowed_operations.includes('section.page.patch')
  const value = pagePatch(section)
  const landscape = value.orientation === 'landscape'
  const paper = papers.find(item => value.width_twips === (landscape ? item.height : item.width) && value.height_twips === (landscape ? item.width : item.height))
  const margins = marginPresets.find(item => item.top === value.margin_top_twips && item.bottom === value.margin_bottom_twips && item.left === value.margin_left_twips && item.right === value.margin_right_twips)
  return <>
    <LayoutMenu icon="margins" label="Margins" disabled={readOnly} value={margins?.name ?? 'Custom'} items={marginPresets.map(preset => ({
      id: preset.name, label: preset.name, detail: `Top ${inches(preset.top)} Bottom ${inches(preset.bottom)} Left ${inches(preset.left)} Right ${inches(preset.right)}`,
      selected: margins?.name === preset.name,
      apply: () => onChange({ ...value, margin_top_twips: preset.top, margin_bottom_twips: preset.bottom, margin_left_twips: preset.left, margin_right_twips: preset.right }),
    }))} />
    <LayoutMenu icon="orientation" label="Orientation" disabled={readOnly} value={landscape ? 'Landscape' : 'Portrait'} items={(['portrait', 'landscape'] as const).map(orientation => ({
      id: orientation, label: orientation === 'portrait' ? 'Portrait' : 'Landscape', selected: value.orientation === orientation,
      apply: () => { if (orientation !== value.orientation) onChange({ ...value, orientation, width_twips: value.height_twips, height_twips: value.width_twips }) },
    }))} />
    <LayoutMenu icon="pageSetup" label="Size" disabled={readOnly} value={paper?.name ?? `${inches(value.width_twips)} × ${inches(value.height_twips)}`} items={papers.map(item => ({
      id: item.name, label: item.name, detail: `${inches(item.width)} × ${inches(item.height)}`, selected: paper?.name === item.name,
      apply: () => onChange({ ...value, width_twips: landscape ? item.height : item.width, height_twips: landscape ? item.width : item.height }),
    }))} />
  </>
}

/** A compact spinner with its label beside it, as in Word's Indent and Spacing fields. */
function LayoutField({ label, unit, value, step, min, max, disabled, onChange }: { label: string; unit: string; value?: number; step: number; min: number; max: number; disabled: boolean; onChange(value: number | null): void }) {
  const shown = value === undefined ? '' : String(Number(value.toFixed(2)))
  const [draft, setDraft] = useState(shown)
  useEffect(() => setDraft(shown), [shown])
  const amount = Number(draft), valid = draft.trim() === '' || (Number.isFinite(amount) && amount >= min && amount <= max)
  function commit() {
    if (!valid) { setDraft(shown); return }
    const next = draft.trim() === '' ? null : amount
    if (next !== (value ?? null)) onChange(next)
  }
  return <label className="document-layout-field">
    <span>{label}</span>
    <input type="number" aria-label={`${label} in ${unit}`} title={`${label} in ${unit}`} placeholder="–" value={draft} disabled={disabled} min={min} max={max} step={step}
      aria-invalid={!valid || undefined}
      onChange={event => setDraft(event.target.value)}
      onBlur={commit}
      onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur(); if (event.key === 'Escape') setDraft(shown) }} />
  </label>
}

const toInches = (value?: number) => value === undefined ? undefined : value / 1440
const toPoints = (value?: number) => value === undefined ? undefined : value / 20
const fromInches = (value: number | null) => value === null ? null : Math.round(value * 1440)
const fromPoints = (value: number | null) => value === null ? null : Math.round(value * 20)
const lineSpacings = [['', 'Inherited'], ['auto:240', 'Single'], ['auto:276', '1.15 lines'], ['auto:360', '1.5 lines'], ['auto:480', 'Double']] as const

/**
 * Word's Layout → Paragraph: Indent left/right in inches and Spacing before/after in points,
 * with the settings Word keeps in its Paragraph dialog behind the same kind of launcher.
 */
export function DocumentParagraphLayout({ properties, disabled, onChange }: { properties?: LayoutParagraphSettings; disabled: boolean; onChange(patch: LayoutParagraphPatch): void }) {
  const inactive = disabled || !properties
  const [more, setMore] = useState(false)
  const box = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!more) return
    const outside = (event: Event) => { if (!box.current?.contains(event.target as Node)) setMore(false) }
    const escape = (event: KeyboardEvent) => { if (event.key === 'Escape') setMore(false) }
    window.document.addEventListener('mousedown', outside)
    window.document.addEventListener('keydown', escape)
    return () => { window.document.removeEventListener('mousedown', outside); window.document.removeEventListener('keydown', escape) }
  }, [more])
  const line = properties?.line_spacing === undefined ? '' : `${properties.line_rule ?? 'auto'}:${properties.line_spacing}`
  return <div className="document-layout-paragraph" aria-label="Indent and spacing" ref={box}>
    <div>
      <LayoutField label="Indent left" unit="inches" step={0.1} min={-22} max={22} disabled={inactive} value={toInches(properties?.indent_left_twips)} onChange={value => onChange({ indent_left_twips: fromInches(value) })} />
      <LayoutField label="Indent right" unit="inches" step={0.1} min={-22} max={22} disabled={inactive} value={toInches(properties?.indent_right_twips)} onChange={value => onChange({ indent_right_twips: fromInches(value) })} />
    </div>
    <div>
      <LayoutField label="Spacing before" unit="points" step={6} min={0} max={1584} disabled={inactive} value={toPoints(properties?.spacing_before_twips)} onChange={value => onChange({ spacing_before_twips: fromPoints(value) })} />
      <LayoutField label="Spacing after" unit="points" step={6} min={0} max={1584} disabled={inactive} value={toPoints(properties?.spacing_after_twips)} onChange={value => onChange({ spacing_after_twips: fromPoints(value) })} />
    </div>
    <RibbonButton className="document-layout-more" icon="more" label="Paragraph settings" labelHidden title="Line spacing, first line, hanging and outline level" disabled={inactive} aria-expanded={more} aria-haspopup="dialog" onClick={() => setMore(value => !value)} />
    {more && <div className="document-layout-panel" role="group" aria-label="Paragraph settings">
      <LayoutField label="First line" unit="inches" step={0.1} min={0} max={22} disabled={inactive} value={toInches(properties?.first_line_twips)} onChange={value => onChange({ first_line_twips: fromInches(value), hanging_twips: null })} />
      <LayoutField label="Hanging" unit="inches" step={0.1} min={0} max={22} disabled={inactive} value={toInches(properties?.hanging_twips)} onChange={value => onChange({ hanging_twips: fromInches(value), first_line_twips: null })} />
      <label><span>Line spacing</span><select aria-label="Paragraph line spacing" disabled={inactive} value={lineSpacings.some(([id]) => id === line) ? line : 'custom'} onChange={event => { const next = event.target.value; onChange(next ? { line_rule: 'auto', line_spacing: Number(next.split(':')[1]) } : { line_rule: null, line_spacing: null }) }}>
        {lineSpacings.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
        {!lineSpacings.some(([id]) => id === line) && <option value="custom">Custom ({properties?.line_rule && properties.line_rule !== 'auto' ? `${(properties.line_spacing ?? 0) / 20} pt` : `${(properties?.line_spacing ?? 0) / 240} lines`})</option>}
      </select></label>
      <label><span>Outline level</span><select aria-label="Paragraph outline level" disabled title="Outline level changes are not supported by the DOCX engine" value={properties?.outline_level ?? ''} onChange={event => onChange({ outline_level: event.target.value === '' ? null : Number(event.target.value) })}>
        <option value="">From style</option><option value="9">Body text</option>
        {Array.from({ length: 9 }, (_, level) => <option key={level} value={level}>Level {level + 1}</option>)}
      </select></label>
    </div>}
  </div>
}
