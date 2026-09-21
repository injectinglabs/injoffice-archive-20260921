/**
 * Where a click that missed the text puts the caret. Word never ignores a click on the page: a
 * click in the empty part of a line lands in that line, and a click below the last line lands at
 * the end of the last paragraph. The geometry is kept pure so it can be tested without a DOM.
 */
export interface CaretBox { left: number; right: number; top: number; bottom: number }
/** The editable box the click belongs to, and a point inside it for `caretRangeFromPoint`. */
export interface CaretTarget { index: number; x: number; y: number }

/** Lines rarely align to the pixel; boxes within this distance count as the same line. */
const LINE_TOLERANCE = 1
function clamp(value: number, low: number, high: number) { return low > high ? (low + high) / 2 : Math.min(Math.max(value, low), high) }
function inside(box: CaretBox, x: number, y: number): { x: number; y: number } {
  return { x: clamp(x, box.left + .5, box.right - .5), y: clamp(y, box.top + .5, box.bottom - .5) }
}

export function caretTargetAt(boxes: readonly CaretBox[], x: number, y: number): CaretTarget | undefined {
  if (!boxes.length) return
  let last = 0
  boxes.forEach((box, index) => { if (box.bottom >= boxes[last]!.bottom) last = index })
  // Below every line: the end of the last paragraph, as Word does.
  if (y > boxes[last]!.bottom) return { index: last, ...inside(boxes[last]!, boxes[last]!.right, (boxes[last]!.top + boxes[last]!.bottom) / 2) }
  let best: { index: number; dy: number; dx: number } | undefined
  boxes.forEach((box, index) => {
    const dy = Math.max(box.top - y, y - box.bottom, 0), dx = Math.max(box.left - x, x - box.right, 0)
    if (!best) { best = { index, dy, dx }; return }
    if (Math.abs(dy - best.dy) <= LINE_TOLERANCE ? dx < best.dx : dy < best.dy) best = { index, dy: Math.min(dy, best.dy), dx }
  })
  return { index: best!.index, ...inside(boxes[best!.index]!, x, y) }
}
