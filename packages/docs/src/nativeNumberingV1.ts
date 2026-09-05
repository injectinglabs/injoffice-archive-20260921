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

/** Places a measured marker at the numbering text-margin anchor. */
export function positionNativeDocxListMarkerV1(numbering: NativeDocxResolvedNumberingV1, direction: 'ltr' | 'rtl', markerAdvanceMilliPoints: number): NativeDocxListMarkerGeometryV1 | undefined {
  const labelStart = twipsToMilliPoints(numbering.label_start_twips)
  const labelEnd = twipsToMilliPoints(numbering.label_end_twips)
  const bodyTextStart = twipsToMilliPoints(numbering.text_start_twips)
  if (labelStart === undefined || labelEnd === undefined || bodyTextStart === undefined || !Number.isSafeInteger(markerAdvanceMilliPoints) || markerAdvanceMilliPoints <= 0 || labelStart < 0 || labelEnd <= labelStart || bodyTextStart !== labelEnd) return undefined
  const leading = numbering.alignment === 'start' ? direction === 'ltr' : numbering.alignment === 'end' ? direction === 'rtl' : numbering.alignment === 'left'
  const trailing = numbering.alignment === 'end' ? direction === 'ltr' : numbering.alignment === 'start' ? direction === 'rtl' : numbering.alignment === 'right'
  const markerStart = leading ? labelStart : trailing ? labelStart - markerAdvanceMilliPoints : labelStart - Math.round(markerAdvanceMilliPoints / 2)
  const markerEnd = markerStart + markerAdvanceMilliPoints
  if (!Number.isSafeInteger(markerStart) || markerStart < 0 || !Number.isSafeInteger(markerEnd) || markerEnd > labelEnd) return undefined
  return { label_start_millipoints: labelStart, label_end_millipoints: labelEnd, body_text_start_millipoints: bodyTextStart, marker_start_millipoints: markerStart }
}

/** Resolves the exact implicit hanging stop or the next bounded default stop. */
export function nativeDocxListSuffixTabTargetV1(currentMilliPoints: number, bodyTextStartMilliPoints: number, defaultIntervalMilliPoints: number, numberingTabMilliPoints?: number): number | undefined {
  if (![currentMilliPoints, bodyTextStartMilliPoints, defaultIntervalMilliPoints].every(Number.isSafeInteger) || currentMilliPoints < 0 || bodyTextStartMilliPoints < 0 || defaultIntervalMilliPoints <= 0) return undefined
  if (numberingTabMilliPoints !== undefined) return Number.isSafeInteger(numberingTabMilliPoints) && numberingTabMilliPoints > currentMilliPoints ? numberingTabMilliPoints : undefined
  const nextDefault = (Math.floor(currentMilliPoints / defaultIntervalMilliPoints) + 1) * defaultIntervalMilliPoints
  const target = currentMilliPoints < bodyTextStartMilliPoints ? bodyTextStartMilliPoints : nextDefault
  return Number.isSafeInteger(target) && target > currentMilliPoints ? target : undefined
}
