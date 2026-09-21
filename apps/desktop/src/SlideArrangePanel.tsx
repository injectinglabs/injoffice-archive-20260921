import type { NativeElement } from '@injoffice/pptx-native';
import { elementKey } from './presentationCommands';
import type { ArrangeAction } from './presentationArrange';
import { shortcutPlatform } from './shortcuts';
const actions: [ArrangeAction, string][] = [['left','Align left'],['center','Align center'],['right','Align right'],['top','Align top'],['middle','Align middle'],['bottom','Align bottom'],['horizontal','Space horizontally'],['vertical','Space vertically']];
/**
 * PowerPoint's Home › Drawing › Arrange, as a labelled group inside the Format
 * pane. It is a plain section, not a disclosure link: the pane itself is what
 * opens and closes.
 */
export default function SlideArrangePanel({elements,keys,disabled,onToggle,onArrange}: {elements: NativeElement[];keys: string[];disabled:boolean;onToggle(key:string):void;onArrange(action:ArrangeAction):void}) {
  return <section className="presentation-arrange" aria-label="Arrange objects">
    <h3>Arrange{keys.length > 1 ? ` · ${keys.length} selected` : ''}</h3>
    <p className="presentation-help">{`${shortcutPlatform() === 'mac' ? '⌘-click' : 'Ctrl+click'} objects on the slide, or tick them below. Alignment uses unrotated layout boxes; rotations stay unchanged. Spacing keeps the outer objects fixed.`}</p>
    <fieldset disabled={disabled} aria-label="Objects to arrange"><div className="presentation-arrange-list">{elements.map((element,index) => <label key={elementKey(element)}><input type="checkbox" checked={keys.includes(elementKey(element))} onChange={() => onToggle(elementKey(element))} />{element.name || `${element.kind} ${index + 1}`}</label>)}</div></fieldset>
    <div className="presentation-arrange-actions">{actions.map(([action,label]) => <button key={action} disabled={disabled || keys.length < (action === 'horizontal' || action === 'vertical' ? 3 : 2)} onClick={() => onArrange(action)}>{label}</button>)}</div>
    {!elements.length && <p className="presentation-help">This slide has no editable top-level text, shapes or pictures.</p>}
  </section>;
}
