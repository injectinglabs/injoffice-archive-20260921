// Derived from bidi-js 1.1.0, Copyright (c) 2021 Jason Johnston. MIT.
// License: packages/pdf/third-party/BIDI-JS-LICENSE.txt.
// Local patch: trailing whitespace reset uses slice-relative indices.
// Local boundary supplies Unicode scalars as an indexed array, not UTF-16 units.
import { getBidiCharType, TRAILING_TYPES } from './charTypes.js'

/**
 * Given a start and end denoting a single line within a string, and a set of precalculated
 * bidi embedding levels, produce a list of segments whose ordering should be flipped, in sequence.
 * @param {string} string - the full input string
 * @param {GetEmbeddingLevelsResult} embeddingLevelsResult - the result object from getEmbeddingLevels
 * @param {number} [start] - first character in a subset of the full string
 * @param {number} [end] - last character in a subset of the full string
 * @return {number[][]} - the list of start/end segments that should be flipped, in order.
 */
export function getReorderSegments(string, embeddingLevelsResult, start, end) {
  let strLen = string.length
  start = Math.max(0, start == null ? 0 : +start)
  end = Math.min(strLen - 1, end == null ? strLen - 1 : +end)

  const segments = []
  embeddingLevelsResult.paragraphs.forEach(paragraph => {
    const lineStart = Math.max(start, paragraph.start)
    const lineEnd = Math.min(end, paragraph.end)
    if (lineStart < lineEnd) {
      // Local slice for mutation
      const lineLevels = embeddingLevelsResult.levels.slice(lineStart, lineEnd + 1)

      // 3.4 L1.4: Reset any sequence of whitespace characters and/or isolate formatting characters at the
      // end of the line to the paragraph level.
      for (let i = lineEnd; i >= lineStart && (getBidiCharType(string[i]) & TRAILING_TYPES); i--) {
        lineLevels[i - lineStart] = paragraph.level
      }

      // L2. From the highest level found in the text to the lowest odd level on each line, including intermediate levels
      // not actually present in the text, reverse any contiguous sequence of characters that are at that level or higher.
      let maxLevel = paragraph.level
      let minOddLevel = Infinity
      for (let i = 0; i < lineLevels.length; i++) {
        const level = lineLevels[i]
        if (level > maxLevel) maxLevel = level
        if (level < minOddLevel) minOddLevel = level | 1
      }
      for (let lvl = maxLevel; lvl >= minOddLevel; lvl--) {
        for (let i = 0; i < lineLevels.length; i++) {
          if (lineLevels[i] >= lvl) {
            const segStart = i
            while (i + 1 < lineLevels.length && lineLevels[i + 1] >= lvl) {
              i++
            }
            if (i > segStart) {
              segments.push([segStart + lineStart, i + lineStart])
            }
          }
        }
      }
    }
  })
  return segments
}

// String rewriting/mirroring is intentionally omitted: HarfBuzz paints the
// original text with resolved direction. This module only orders scalar indices.

/**
 * @param {string} string
 * @param {GetEmbeddingLevelsResult} embedLevelsResult
 * @param {number} [start]
 * @param {number} [end]
 * @return {number[]} an array with character indices in their new bidi order
 */
export function getReorderedIndices(string, embedLevelsResult, start, end) {
  const segments = getReorderSegments(string, embedLevelsResult, start, end)
  // Fill an array with indices
  const indices = []
  for (let i = 0; i < string.length; i++) {
    indices[i] = i
  }
  // Reverse each segment in order
  segments.forEach(([start, end]) => {
    const slice = indices.slice(start, end + 1)
    for (let i = slice.length; i--;) {
      indices[end - i] = slice[i]
    }
  })
  return indices
}
