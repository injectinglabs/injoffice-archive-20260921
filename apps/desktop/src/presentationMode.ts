import type { NativePptxDeck } from '@injoffice/pptx-native';

export interface PresentationModeState { readonly deck: NativePptxDeck; readonly index: number; readonly hasDraft: boolean }
/** Own navigation state; never reuse the editor's selection or history cursor. */
export function startPresentationMode(deck: NativePptxDeck, index: number, hasDraft: boolean): PresentationModeState {
  if (!deck.slides.length) throw new Error('The presentation has no slides.');
  return { deck, index: Math.max(0, Math.min(deck.slides.length - 1, Math.trunc(Number.isFinite(index) ? index : 0))), hasDraft };
}
export function navigatePresentation(state: PresentationModeState, key: string): PresentationModeState | null {
  if (key === 'Escape') return null;
  const last = state.deck.slides.length - 1;
  const next = key === 'Home' ? 0 : key === 'End' ? last : ['ArrowRight', 'ArrowDown', 'PageDown', ' '].includes(key) ? Math.min(last, state.index + 1) : ['ArrowLeft', 'ArrowUp', 'PageUp'].includes(key) ? Math.max(0, state.index - 1) : state.index;
  return next === state.index ? state : { ...state, index: next };
}
export function presentationScale(width: number, height: number, slideWidth: number, slideHeight: number): number {
  if (![width,height,slideWidth,slideHeight].every(Number.isFinite) || width <= 0 || height <= 0 || slideWidth <= 0 || slideHeight <= 0) return 0;
  return Math.min(width / slideWidth, height / slideHeight);
}
