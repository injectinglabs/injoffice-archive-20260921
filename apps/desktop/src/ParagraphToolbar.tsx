import { useEffect, useState } from 'react';
import type { NativeDocxDocumentV1, NativeDocxNumberingReferenceV1 } from '../../../packages/docs/src/nativeContract';
import type { NativeDocxParagraphPropertyPatchV1 } from '../../../packages/docs/src/nativeTransactionAdapterV1';
export interface ParagraphSettings {
  numbering?: NativeDocxNumberingReferenceV1; paragraph_style_id?: string; outline_level?: number; spacing_before_twips?: number; spacing_after_twips?: number;
  line_spacing?: number; line_rule?: 'auto' | 'exact' | 'atLeast'; indent_left_twips?: number; indent_right_twips?: number; first_line_twips?: number; hanging_twips?: number;
}
export type ParagraphPatch = NativeDocxParagraphPropertyPatchV1;
function PointsField({ label, value, disabled, signed, onChange }: { label: string; value?: number; disabled: boolean; signed?: boolean; onChange(value: number | null): void }) {
  const [draft, setDraft] = useState(value === undefined ? '' : String(value / 20));
  useEffect(() => setDraft(value === undefined ? '' : String(value / 20)), [value]);
  const amount = Number(draft), valid = draft.trim() === '' || (Number.isFinite(amount) && amount >= (signed ? -1584 : 0) && amount <= 1584);
  function commit() { if (!valid) return; const next = draft.trim() === '' ? null : Math.round(amount * 20); if (next !== (value ?? null)) onChange(next); }
  return <label>{label}<input type="number" aria-label={`${label} in points`} value={draft} disabled={disabled} min={signed ? -1584 : 0} max={1584} step="0.5" aria-invalid={!valid || undefined} onChange={event => setDraft(event.target.value)} onBlur={commit} onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur(); if (event.key === 'Escape') setDraft(value === undefined ? '' : String(value / 20)); }} /></label>;
}
export default function ParagraphToolbar({ properties, styles, numbering = [], section='all', disabled, onChange }: { section?:'all'|'home'|'layout'|'styles'|'list'; properties?: ParagraphSettings; styles: Array<{id:string;name:string}>; numbering?: NativeDocxDocumentV1['numbering_definitions']; disabled: boolean; onChange(patch: ParagraphPatch): void }) {
  const inactive = disabled || !properties;
  const line = properties?.line_spacing === undefined ? '' : `${properties.line_rule ?? 'auto'}:${properties.line_spacing}`;
  const list = numbering.find(item => item.num_id === properties?.numbering?.num_id);
  const commonLines = ['', 'auto:240','auto:276','auto:360','auto:480'];
  const showStyles = section==='all'||section==='home'||section==='styles', showList = section==='all'||section==='home'||section==='list';
  return <div className={`paragraph-toolbar paragraph-toolbar-${section}`} aria-label="Paragraph formatting">
    {showStyles&&<select className="ribbon-combo ribbon-combo-style" aria-label="Paragraph style" title="Paragraph style" disabled={inactive || !styles.length} value={properties?.paragraph_style_id ?? ''} onChange={event => onChange({paragraph_style_id:event.target.value || null,outline_level:null})}><option value="" />{styles.map(style => <option key={style.id} value={style.id}>{style.name}</option>)}{properties?.paragraph_style_id && !styles.some(style => style.id === properties.paragraph_style_id) && <option value={properties.paragraph_style_id}>{properties.paragraph_style_id}</option>}</select>}
    {showList&&<>
    <select className="ribbon-combo ribbon-combo-list" aria-label="Paragraph list" title="Paragraph list" disabled={inactive} value={properties?.numbering?.num_id ?? ''} onChange={event => {const id = event.target.value; onChange({numbering_num_id:id || null,numbering_level:!id ? null : id === '0' ? 0 : numbering.find(item => item.num_id === id)!.levels[0].level});}}><option value="" /><option value="0">No list</option>{numbering.map(item => <option key={item.num_id} value={item.num_id}>{item.levels[0].format === 'bullet' ? 'Bullets' : 'Numbered'} · {item.num_id}</option>)}{properties?.numbering && properties.numbering.num_id !== '0' && !list && <option value={properties.numbering.num_id}>Existing list · {properties.numbering.num_id}</option>}</select>
    {list && <select className="ribbon-combo ribbon-combo-size" aria-label="List level" title="List level" disabled={inactive} value={properties?.numbering?.level ?? list.levels[0].level} onChange={event => onChange({numbering_num_id:list.num_id,numbering_level:Number(event.target.value)})}>{list.levels.map(level => <option key={level.level} value={level.level}>{level.level + 1}</option>)}</select>}
    </>}
    {(section==='all'||section==='layout')&&<>
    <label>Outline<select className="ribbon-combo ribbon-combo-list" aria-label="Paragraph outline level" title="Paragraph outline level" disabled={inactive} value={properties?.outline_level ?? ''} onChange={event => onChange({outline_level:event.target.value === '' ? null : Number(event.target.value)})}><option value="" /><option value="9">Body text</option>{Array.from({length:9},(_,level) => <option key={level} value={level}>Level {level + 1}</option>)}</select></label>
    <label>Line spacing<select className="ribbon-combo ribbon-combo-list" aria-label="Paragraph line spacing" title="Paragraph line spacing" disabled={inactive} value={line} onChange={event => { const value = event.target.value; onChange(value ? {line_rule:'auto',line_spacing:Number(value.split(':')[1])} : {line_rule:null,line_spacing:null}); }}><option value="" /><option value="auto:240">Single</option><option value="auto:276">1.15 lines</option><option value="auto:360">1.5 lines</option><option value="auto:480">Double</option>{!commonLines.includes(line) && <option value={line}>Custom ({properties?.line_rule === 'auto' || !properties?.line_rule ? `${(properties?.line_spacing ?? 0) / 240} lines` : `${(properties?.line_spacing ?? 0) / 20} pt`})</option>}</select></label>
    <PointsField label="Before" value={properties?.spacing_before_twips} disabled={inactive} onChange={value => onChange({spacing_before_twips:value})} />
    <PointsField label="After" value={properties?.spacing_after_twips} disabled={inactive} onChange={value => onChange({spacing_after_twips:value})} />
    <PointsField label="Left indent" value={properties?.indent_left_twips} disabled={inactive} signed onChange={value => onChange({indent_left_twips:value})} />
    <PointsField label="Right indent" value={properties?.indent_right_twips} disabled={inactive} signed onChange={value => onChange({indent_right_twips:value})} />
    <PointsField label="First line" value={properties?.first_line_twips} disabled={inactive} onChange={value => onChange({first_line_twips:value,hanging_twips:null})} />
    <PointsField label="Hanging" value={properties?.hanging_twips} disabled={inactive} onChange={value => onChange({hanging_twips:value,first_line_twips:null})} />
    <span>Spacing and indentation in points</span>
    </>}
  </div>;
}
