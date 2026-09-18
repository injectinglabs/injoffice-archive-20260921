/** Pure integer geometry helpers for already-resolved native DOCX markers. */

import { DOCX_MAX_TWIPS_FOR_MILLIPOINTS } from './nativeContract.js'
import type { NativeDocxResolvedNumberingV1 } from './nativeResolvedLayout.js'

export interface NativeDocxListMarkerGeometryV1 {
  label_start_millipoints: number
  label_end_millipoints: number
  body_text_start_millipoints: number
  marker_start_millipoints: number
}

function twipsToMilliPoints(twips: number): number | undefined {
  if (!Number.isSafeInteger(twips) || Object.is(twips, -0) || Math.abs(twips) > DOCX_MAX_TWIPS_FOR_MILLIPOINTS) return undefined
  return twips * 50
}

/** Places a measured marker at the numbering text-margin anchor.
 *
 * A marker wider than its label region is not a refusal and does not push the
 * marker off its anchor: Word draws the whole marker from the anchor and lets the
 * numbering suffix carry the text clear of it (`w:suff w:val="tab"`, the default,
 * advances to the next stop past the marker end). Verified against Word's own
 * PDF export of `cjklist34/35/44.docx`, whose `w:ind w:left="480" w:hanging="480"`
 * gives a 24 pt label region that a three-glyph CJK marker (`壹拾壹.`, 39.029 pt)
 * overruns by 15 pt while Word still starts the marker at the anchor. Only a
 * marker that would start left of the text origin is still refused. */
export function positionNativeDocxListMarkerV1(numbering: NativeDocxResolvedNumberingV1, direction: 'ltr' | 'rtl', markerAdvanceMilliPoints: number): NativeDocxListMarkerGeometryV1 | undefined {
  const labelStart = twipsToMilliPoints(numbering.label_start_twips)
  const labelEnd = twipsToMilliPoints(numbering.label_end_twips)
  const bodyTextStart = twipsToMilliPoints(numbering.text_start_twips)
  if (labelStart === undefined || labelEnd === undefined || bodyTextStart === undefined || !Number.isSafeInteger(markerAdvanceMilliPoints) || markerAdvanceMilliPoints <= 0 || labelStart < 0 || labelEnd <= labelStart || bodyTextStart !== labelEnd) return undefined
  const leading = numbering.alignment === 'start' ? direction === 'ltr' : numbering.alignment === 'end' ? direction === 'rtl' : numbering.alignment === 'left'
  const trailing = numbering.alignment === 'end' ? direction === 'ltr' : numbering.alignment === 'start' ? direction === 'rtl' : numbering.alignment === 'right'
  const markerStart = leading ? labelStart : trailing ? labelStart - markerAdvanceMilliPoints : labelStart - Math.round(markerAdvanceMilliPoints / 2)
  const markerEnd = markerStart + markerAdvanceMilliPoints
  if (!Number.isSafeInteger(markerStart) || markerStart < 0 || !Number.isSafeInteger(markerEnd)) return undefined
  return { label_start_millipoints: labelStart, label_end_millipoints: labelEnd, body_text_start_millipoints: bodyTextStart, marker_start_millipoints: markerStart }
}

/** Resolves the exact implicit hanging stop or the next bounded default stop.
 *
 * The hanging indent's implicit stop only applies while the marker still ends
 * before it; once the marker has run past it the tab falls back to the default
 * grid, which is measured from the text margin, not from the marker. Word's PDF
 * for `cjklist34.docx` pins both arms: item 10 (`壹拾.`, marker end 27.029 pt) and
 * item 11 (`壹拾壹.`, marker end 39.029 pt) both start their text at 48 pt past the
 * margin, the same default stop, not at marker end plus a fixed gap. */
export function nativeDocxListSuffixTabTargetV1(currentMilliPoints: number, bodyTextStartMilliPoints: number, defaultIntervalMilliPoints: number, numberingTabMilliPoints?: number): number | undefined {
  if (![currentMilliPoints, bodyTextStartMilliPoints, defaultIntervalMilliPoints].every(Number.isSafeInteger) || currentMilliPoints < 0 || bodyTextStartMilliPoints < 0 || defaultIntervalMilliPoints <= 0) return undefined
  if (numberingTabMilliPoints !== undefined) return Number.isSafeInteger(numberingTabMilliPoints) && numberingTabMilliPoints > currentMilliPoints ? numberingTabMilliPoints : undefined
  const nextDefault = (Math.floor(currentMilliPoints / defaultIntervalMilliPoints) + 1) * defaultIntervalMilliPoints
  const target = currentMilliPoints < bodyTextStartMilliPoints ? bodyTextStartMilliPoints : nextDefault
  return Number.isSafeInteger(target) && target > currentMilliPoints ? target : undefined
}
