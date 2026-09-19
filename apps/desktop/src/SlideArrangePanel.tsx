import type { NativeElement } from '@injoffice/pptx-native';
import { elementKey } from './presentationCommands';
import type { ArrangeAction } from './presentationArrange';
const actions: [ArrangeAction, string][] = [['left','Align left'],['center','Align center'],['right','Align right'],['top','Align top'],['middle','Align middle'],['bottom','Align bottom'],['horizontal','Space horizontally'],['vertical','Space vertically']];
export default function SlideArrangePanel({elements,keys,disabled,onToggle,onArrange}: {elements: NativeElement[];keys: string[];disabled:boolean;onToggle(key:string):void;onArrange(action:ArrangeAction):void}) {
  return <details className="presentation-arrange" open={keys.length > 1 || undefined}>
    <summary>Arrange objects{keys.length ? ` · ${keys.length} selected` : ''}</summary>
    <p className="presentation-help">Ctrl / ⌘ click objects or select below. Alignment uses unrotated layout boxes; rotations stay unchanged. Spacing keeps the outer objects fixed.</p>
    <fieldset disabled={disabled} aria-label="Objects to arrange"><div className="presentation-arrange-list">{elements.map((element,index) => <label key={elementKey(element)}><input type="checkbox" checked={keys.includes(elementKey(element))} onChange={() => onToggle(elementKey(element))} />{element.name || `${element.kind} ${index + 1}`}</label>)}</div></fieldset>
    <div className="presentation-arrange-actions">{actions.map(([action,label]) => <button key={action} disabled={disabled || keys.length < (action === 'horizontal' || action === 'vertical' ? 3 : 2)} onClick={() => onArrange(action)}>{label}</button>)}</div>
    {!elements.length && <p className="presentation-help">This slide has no editable top-level text, shapes or pictures.</p>}
  </details>;
}
