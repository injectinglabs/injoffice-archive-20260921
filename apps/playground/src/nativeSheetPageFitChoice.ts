/** Explicit host inputs only; blank is not the zero/unlimited choice. */
export function nativeSheetPageFitChoice(width: string, height: string): { width: number; height: number } {
  const valid = (value: string) => /^(0|[1-9][0-9]{0,2})$/.test(value) && Number(value) <= 100
  if (!valid(width) || !valid(height) || (Number(width) === 0 && Number(height) === 0)) {
    throw new Error('Enter whole-number page limits from 0 to 100, with at least one positive limit. Use 0 for an unlimited dimension.')
  }
  return { width: Number(width), height: Number(height) }
}
