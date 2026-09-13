// Derived from bidi-js 1.1.0, Copyright (c) 2021 Jason Johnston. MIT.
// License: packages/pdf/third-party/BIDI-JS-LICENSE.txt.
// Local boundary supplies Unicode scalars as an indexed array, not UTF-16 units.
import { BRACKET_PAIRS, CANONICAL_BRACKETS } from './unicode17.generated.js'
const pairs = new Map(Object.entries(BRACKET_PAIRS))
const reverse = new Map([...pairs].map(([a, b]) => [b, a]))
const canonical = new Map(Object.entries(CANONICAL_BRACKETS))
export const openingToClosingBracket = char => pairs.get(char) || null
export const closingToOpeningBracket = char => reverse.get(char) || null
export const getCanonicalBracket = char => canonical.get(char) || null
