import type { NativeElement, NativePptxDeck } from '@injoffice/pptx-native';
import type { PptxNativeMutationRequestV1 } from '@injoffice/pptx-wasm';
import { elementKey, shapeCommand, shapeTarget, transformCommand, geometryTarget } from './presentationCommands';

export type ArrangeAction = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom' | 'horizontal' | 'vertical';
export function arrangeTargets(deck: NativePptxDeck, index: number): NativeElement[] {
  return (deck.slides[index]?.elements ?? []).filter(element => element.source && element.compatibility.status === 'editable' &&
    (element.kind === 'shape' ? !!shapeTarget(deck, elementKey(element)) :
      (element.kind === 'text' || element.kind === 'picture') && !!geometryTarget(deck, elementKey(element))));
}
export function toggleArrangeSelection(current: readonly string[], key: string): string[] {
  return current.includes(key) ? current.filter(value => value !== key) : [...current, key];
}
/** Align unrotated layout boxes, retaining each object's rotation and appearance. */
export function arrangeCommand(deck: NativePptxDeck, index: number, keys: readonly string[], action: ArrangeAction, operationId: string): PptxNativeMutationRequestV1 {
  const horizontal = action === 'left' || action === 'center' || action === 'right' || action === 'horizontal';
  const distribute = action === 'horizontal' || action === 'vertical';
  if (!['left','center','right','top','middle','bottom','horizontal','vertical'].includes(action)) throw new Error('Unknown arrangement.');
  if (new Set(keys).size !== keys.length) throw new Error('Select each object only once.');
  if (keys.length < (distribute ? 3 : 2) || keys.length > 100) throw new Error(`Select ${distribute ? '3' : '2'} to 100 editable objects on this slide.`);
  const allowed = arrangeTargets(deck, index);
  const elements = keys.map(key => { const element = allowed.find(item => elementKey(item) === key); if (!element) throw new Error('Only exact top-level text, shapes and pictures on this slide can be arranged.'); return element; });
  const axis = horizontal ? 'x' : 'y', extent = horizontal ? 'cx' : 'cy';
  const start = Math.min(...elements.map(item => item.transform[axis]));
  const end = Math.max(...elements.map(item => item.transform[axis] + item.transform[extent]));
  const sorted = [...elements].sort((a,b) => a.transform[axis] - b.transform[axis] || elementKey(a).localeCompare(elementKey(b)));
  // Keep the first and last objects fixed during spacing. Containment can put
  // the rightmost edge on an interior object; refuse rather than swap order.
  const spaceStart = sorted[0]!.transform[axis];
  const last = sorted.at(-1)!;
  const spaceEnd = last.transform[axis] + last.transform[extent];
  const gap = (spaceEnd - spaceStart - sorted.reduce((sum,item) => sum + item.transform[extent],0)) / (elements.length - 1);
  if (distribute && gap < 0) throw new Error('Move the outer objects farther apart before distributing equal gaps.');
  let cursor = spaceStart;
  const operations: PptxNativeMutationRequestV1['operations'][number][] = [];
  for (const element of sorted) {
    const old = element.transform;
    const position = distribute ? Math.round(cursor) : action === 'left' || action === 'top' ? start : action === 'right' || action === 'bottom' ? end - old[extent] : Math.round((start + end - old[extent]) / 2);
    cursor += old[extent] + gap;
    if (position === old[axis]) continue;
    if (!Number.isSafeInteger(position)) throw new Error('Arrangement exceeds supported coordinates.');
    const transform = {...old, [axis]: position};
    const key = elementKey(element);
    const id = `${operationId}-${operations.length}`;
    const request = element.kind === 'shape' ? shapeCommand(deck,key,{...shapeTarget(deck,key)!.autoShape,transform},id) : transformCommand(deck,key,transform,id);
    operations.push(...request.operations);
  }
  if (!operations.length) throw new Error('These objects are already arranged.');
  if (!deck.sourceRevision) throw new Error('Open a source presentation before editing.');
  return { expectedSourceRevision: deck.sourceRevision, operations };
}
