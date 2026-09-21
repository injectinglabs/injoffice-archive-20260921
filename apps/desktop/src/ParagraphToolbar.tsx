import type { NativeDocxDocumentV1, NativeDocxNumberingReferenceV1 } from '../../../packages/docs/src/nativeContract';
import type { NativeDocxParagraphPropertyPatchV1 } from '../../../packages/docs/src/nativeTransactionAdapterV1';
export interface ParagraphSettings {
  numbering?: NativeDocxNumberingReferenceV1; paragraph_style_id?: string; outline_level?: number; spacing_before_twips?: number; spacing_after_twips?: number;
  line_spacing?: number; line_rule?: 'auto' | 'exact' | 'atLeast'; indent_left_twips?: number; indent_right_twips?: number; first_line_twips?: number; hanging_twips?: number;
}
export type ParagraphPatch = NativeDocxParagraphPropertyPatchV1;
export default function ParagraphToolbar({ properties, styles, numbering = [], section='all', disabled, onChange }: { section?:'all'|'home'|'styles'|'list'; properties?: ParagraphSettings; styles: Array<{id:string;name:string}>; numbering?: NativeDocxDocumentV1['numbering_definitions']; disabled: boolean; onChange(patch: ParagraphPatch): void }) {
  const inactive = disabled || !properties;
  const list = numbering.find(item => item.num_id === properties?.numbering?.num_id);
  const showStyles = section==='all'||section==='home'||section==='styles', showList = section==='all'||section==='home'||section==='list';
  return <div className={`paragraph-toolbar paragraph-toolbar-${section}`} aria-label="Paragraph formatting">
    {showStyles&&<select className="ribbon-combo ribbon-combo-style" aria-label="Paragraph style" title="Paragraph style" disabled={inactive || !styles.length} value={properties?.paragraph_style_id ?? ''} onChange={event => onChange({paragraph_style_id:event.target.value || null,outline_level:null})}><option value="" />{styles.map(style => <option key={style.id} value={style.id}>{style.name}</option>)}{properties?.paragraph_style_id && !styles.some(style => style.id === properties.paragraph_style_id) && <option value={properties.paragraph_style_id}>{properties.paragraph_style_id}</option>}</select>}
    {showList&&<>
    <select className="ribbon-combo ribbon-combo-list" aria-label="Paragraph list" title="Paragraph list" disabled={inactive} value={properties?.numbering?.num_id ?? ''} onChange={event => {const id = event.target.value; onChange({numbering_num_id:id || null,numbering_level:!id ? null : id === '0' ? 0 : numbering.find(item => item.num_id === id)!.levels[0].level});}}><option value="" /><option value="0">No list</option>{numbering.map(item => <option key={item.num_id} value={item.num_id}>{item.levels[0].format === 'bullet' ? 'Bullets' : 'Numbered'} · {item.num_id}</option>)}{properties?.numbering && properties.numbering.num_id !== '0' && !list && <option value={properties.numbering.num_id}>Existing list · {properties.numbering.num_id}</option>}</select>
    {list && <select className="ribbon-combo ribbon-combo-size" aria-label="List level" title="List level" disabled={inactive} value={properties?.numbering?.level ?? list.levels[0].level} onChange={event => onChange({numbering_num_id:list.num_id,numbering_level:Number(event.target.value)})}>{list.levels.map(level => <option key={level.level} value={level.level}>{level.level + 1}</option>)}</select>}
    </>}
  </div>;
}
