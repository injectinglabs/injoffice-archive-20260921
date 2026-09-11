// Pure display formatting shared by native paint and browser consumers.
const PRINTABLE_ASCII = /^[\x20-\x7E]+$/

export function isCanonicalDisplayText(value: unknown): value is string {
  if (typeof value !== 'string' || value.length < 1 || value.length > 32_767 || /[\u0000-\u001f\u007f]/.test(value)) return false
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index)!
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const trail = value.charCodeAt(index + 1)
      if (trail === undefined || trail < 0xdc00 || trail > 0xdfff) return false
      index += 1
      continue
    }
    if (unit >= 0xdc00 && unit <= 0xdfff) return false
  }
  return true
}

export type NativeSheetCellDisplayFormatResultV2 =
  | { readonly status: 'ready'; readonly text: string }
  | { readonly status: 'refused'; readonly code: 'CELL_DATE_DISPLAY' | 'CELL_STYLE_UNSUPPORTED'; readonly message: string }

/**
 * Locale-independent display for documented OOXML date/number tokens.
 * Uses producer ISO calendar components as written; never Date, Intl, or host TZ.
 * Number cells with `General` keep their stored lexical. Date serial conversion
 * uses only a projected workbook date1904 flag. Named months/days require an
 * explicit OOXML `[$-…]` locale/calendar that this producer has a table for.
 */
export function formatNativeSheetCellDisplayV2(
  kind: 'number' | 'date',
  lexical: string,
  numberFormat: string | undefined,
  date1904?: boolean,
): NativeSheetCellDisplayFormatResultV2 {
  const format = numberFormat ?? 'General'
  if (kind === 'number') {
    if (format === 'General') return { status: 'ready', text: lexical }
    const numeric = classifyFixedNumberFormat(format)
    if (numeric) {
      const fixed = formatFixedDecimalLexical(lexical, numeric.fractionDigits, numeric.percent ? 2 : 0)
      const text = fixed === undefined ? undefined : decorateFixedNumber(fixed, numeric)
      if (text === undefined) return { status: 'refused', code: 'CELL_STYLE_UNSUPPORTED', message: 'bounded numeric format cannot be applied to this lexical' }
      return { status: 'ready', text }
    }
  }
  const classified = classifyOoxmlDateFormat(format)
  if (classified.status === 'ready') {
    let parts: DateTimeParts | undefined
    if (kind === 'date') parts = parseIsoDateTimeParts(lexical)
    else {
      if (classified.needsDate && date1904 === undefined) return { status: 'refused', code: 'CELL_DATE_DISPLAY', message: 'date display would require locale/number-format invention' }
      parts = excelSerialToParts(lexical, date1904 === true, classified.needsDate)
    }
    const text = parts ? applyOoxmlDateTokens(classified.tokens, parts, classified.locale) : undefined
    if (text === undefined) return { status: 'refused', code: 'CELL_DATE_DISPLAY', message: 'date display would require locale/number-format invention' }
    return { status: 'ready', text }
  }
  if (classified.status === 'unusable' || kind === 'date') return { status: 'refused', code: 'CELL_DATE_DISPLAY', message: 'date display would require locale/number-format invention' }
  return { status: 'refused', code: 'CELL_STYLE_UNSUPPORTED', message: 'number format is outside the qualified display subset; the original value remains preserved' }
}

type FixedNumberFormat = { fractionDigits: number; grouped: boolean; percent: boolean; prefix: string; suffix: string }

/** Deliberately one section: no colors, conditions, locale selection, accounting
 * padding, fill characters, scaling commas, fractions or exponent formatting.
 * Only explicit currency literals are accepted; no host currency is inferred. */
function classifyFixedNumberFormat(format: string): FixedNumberFormat | undefined {
  if (format.length > 64) return undefined
  const match = /^(?:"([$£€¥] ?)")?(0|#,##0)(?:\.(0{1,6}))?(%)?(?:"( ?[$£€¥])")?$/.exec(format)
  if (!match || (match[1] && match[5]) || (match[4] && (match[1] || match[5]))) return undefined
  return { prefix: match[1] ?? '', suffix: match[5] ?? '', grouped: match[2] === '#,##0', fractionDigits: match[3]?.length ?? 0, percent: match[4] === '%' }
}

function decorateFixedNumber(fixed: string, format: FixedNumberFormat): string {
  const negative = fixed.startsWith('-')
  const [whole, fraction] = (negative ? fixed.slice(1) : fixed).split('.')
  const grouped = format.grouped ? whole!.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : whole!
  return `${negative ? '-' : ''}${format.prefix}${grouped}${fraction === undefined ? '' : `.${fraction}`}${format.percent ? '%' : ''}${format.suffix}`
}

const EN_US_GREGORIAN_MONTHS_FULL = Object.freeze(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'])
const EN_US_GREGORIAN_MONTHS_ABBR = Object.freeze(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
const EN_US_GREGORIAN_MONTHS_LETTER = Object.freeze(['J', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D'])
const EN_US_GREGORIAN_DAYS_FULL = Object.freeze(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'])
const EN_US_GREGORIAN_DAYS_ABBR = Object.freeze(['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'])

type DateTimeParts = {
  year: number
  month: number
  day: number
  hour: number
  minute: number
  second: number
  weekday: number
  dateValid: boolean
}

type DateLocaleTable = {
  monthsFull: readonly string[]
  monthsAbbr: readonly string[]
  monthsLetter: readonly string[]
  daysFull: readonly string[]
  daysAbbr: readonly string[]
}

const EN_US_GREGORIAN: DateLocaleTable = {
  monthsFull: EN_US_GREGORIAN_MONTHS_FULL,
  monthsAbbr: EN_US_GREGORIAN_MONTHS_ABBR,
  monthsLetter: EN_US_GREGORIAN_MONTHS_LETTER,
  daysFull: EN_US_GREGORIAN_DAYS_FULL,
  daysAbbr: EN_US_GREGORIAN_DAYS_ABBR,
}

type DateFormatToken =
  | { kind: 'year'; width: number }
  | { kind: 'month'; width: number }
  | { kind: 'monthName'; width: number }
  | { kind: 'day'; width: number }
  | { kind: 'weekday'; width: number }
  | { kind: 'hour'; width: number }
  | { kind: 'minute'; width: number }
  | { kind: 'second'; width: number }
  | { kind: 'ampm'; form: 'AM/PM' | 'A/P' }
  | { kind: 'literal'; text: string }

export function formatUsesDateTimeTokens(format: string): boolean {
  return classifyOoxmlDateFormat(format).status !== 'not-date'
}

function classifyOoxmlDateFormat(format: string):
  | { status: 'ready'; tokens: DateFormatToken[]; locale: DateLocaleTable | undefined; needsDate: boolean }
  | { status: 'unusable' }
  | { status: 'not-date' } {
  let body = format
  let locale: DateLocaleTable | undefined
  if (format.startsWith('[$-')) {
    const close = format.indexOf(']')
    if (close < 0) return { status: 'unusable' }
    const inner = format.slice(3, close)
    if (isSystemDateLocale(inner)) return { status: 'unusable' }
    locale = localeTableForOoxmlPrefix(inner)
    body = format.slice(close + 1)
  }
  if (body.includes(';') || body.includes('*') || body.includes('_') || body.includes('[')) {
    return /[ymdhs]/i.test(body) || format.startsWith('[$-') ? { status: 'unusable' } : { status: 'not-date' }
  }
  const tokens = tokenizeOoxmlDateFormat(body)
  if (!tokens) return /[ymdhs]/i.test(body) || format.startsWith('[$-') ? { status: 'unusable' } : { status: 'not-date' }
  const needsNames = tokens.some((token) => token.kind === 'monthName' || token.kind === 'weekday')
  if (needsNames && locale === undefined) return { status: 'unusable' }
  const needsDate = tokens.some((token) => token.kind === 'year' || token.kind === 'month' || token.kind === 'monthName' || token.kind === 'day' || token.kind === 'weekday')
  const needsTime = tokens.some((token) => token.kind === 'hour' || token.kind === 'minute' || token.kind === 'second' || token.kind === 'ampm')
  if (!needsDate && !needsTime) return { status: 'not-date' }
  return { status: 'ready', tokens, locale, needsDate }
}

function isSystemDateLocale(inner: string): boolean {
  return /f800|f400|sysdate|systime/i.test(inner)
}

function localeTableForOoxmlPrefix(inner: string): DateLocaleTable | undefined {
  const head = inner.split(',')[0]!.trim()
  if (/^en[-_]US$/i.test(head)) return EN_US_GREGORIAN
  if (!/^[0-9A-Fa-f]+$/.test(head)) return undefined
  const value = Number.parseInt(head, 16)
  if (!Number.isSafeInteger(value) || value < 0) return undefined
  const lcid = value & 0xffff
  const calendar = (value >>> 16) & 0xff
  if (lcid === 0xf800 || lcid === 0xf400) return undefined
  if (calendar !== 0 && calendar !== 1) return undefined
  if (calendar === 1 || lcid === 0x0409 || lcid === 0x0009) return EN_US_GREGORIAN
  return undefined
}

function tokenizeOoxmlDateFormat(format: string): DateFormatToken[] | undefined {
  const tokens: DateFormatToken[] = []
  let index = 0
  while (index < format.length) {
    const char = format[index]!
    if (char === '[') return undefined
    if (char === '"') {
      const close = format.indexOf('"', index + 1)
      if (close < 0) return undefined
      tokens.push({ kind: 'literal', text: format.slice(index + 1, close) })
      index = close + 1
      continue
    }
    if (char === '\\') {
      if (index + 1 >= format.length) return undefined
      tokens.push({ kind: 'literal', text: format[index + 1]! })
      index += 2
      continue
    }
    const rest = format.slice(index)
    if (/^AM\/PM/i.test(rest) || /^A\/P/i.test(rest)) {
      const ampm = /^AM\/PM/i.test(rest)
      tokens.push({ kind: 'ampm', form: ampm ? 'AM/PM' : 'A/P' })
      index += ampm ? 5 : 3
      continue
    }
    if (char === 'y' || char === 'Y') {
      const width = takeSame(format, index, 'y')
      tokens.push({ kind: 'year', width })
      index += width
      continue
    }
    if (char === 'm' || char === 'M') {
      const width = takeSame(format, index, 'm')
      if (width >= 3) tokens.push({ kind: 'monthName', width })
      else tokens.push({ kind: 'month', width })
      index += width
      continue
    }
    if (char === 'd' || char === 'D') {
      const width = takeSame(format, index, 'd')
      if (width >= 3) tokens.push({ kind: 'weekday', width })
      else tokens.push({ kind: 'day', width })
      index += width
      continue
    }
    if (char === 'h' || char === 'H') {
      const width = takeSame(format, index, 'h')
      tokens.push({ kind: 'hour', width })
      index += width
      continue
    }
    if (char === 's' || char === 'S') {
      const width = takeSame(format, index, 's')
      tokens.push({ kind: 'second', width })
      index += width
      continue
    }
    if (/[A-Za-z]/.test(char)) return undefined
    tokens.push({ kind: 'literal', text: char })
    index += 1
  }
  return resolveMinuteTokens(tokens)
}

function takeSame(format: string, start: number, letter: string): number {
  let width = 0
  while (start + width < format.length && format[start + width]!.toLowerCase() === letter) width += 1
  return width
}

function resolveMinuteTokens(tokens: DateFormatToken[]): DateFormatToken[] {
  const significant = (token: DateFormatToken): boolean => token.kind !== 'literal'
  return tokens.map((token, index) => {
    if (token.kind !== 'month' || token.width > 2) return token
    const prev = tokens.slice(0, index).reverse().find(significant)
    const next = tokens.slice(index + 1).find(significant)
    if (prev?.kind === 'hour' || next?.kind === 'second') return { kind: 'minute', width: token.width }
    return token
  })
}

function applyOoxmlDateTokens(tokens: DateFormatToken[], parts: DateTimeParts, locale: DateLocaleTable | undefined): string | undefined {
  const hour12 = tokens.some((token) => token.kind === 'ampm')
  let text = ''
  for (const token of tokens) {
    switch (token.kind) {
      case 'year':
        text += token.width <= 2 ? String(parts.year % 100).padStart(2, '0') : String(parts.year).padStart(4, '0')
        break
      case 'month':
        text += padDateNumber(parts.month, token.width)
        break
      case 'monthName': {
        if (!locale) return undefined
        const name = token.width >= 5 ? locale.monthsLetter[parts.month - 1] : token.width === 3 ? locale.monthsAbbr[parts.month - 1] : locale.monthsFull[parts.month - 1]
        if (!name) return undefined
        text += name
        break
      }
      case 'day':
        text += padDateNumber(parts.day, token.width)
        break
      case 'weekday': {
        if (!locale) return undefined
        const name = token.width === 3 ? locale.daysAbbr[parts.weekday] : locale.daysFull[parts.weekday]
        if (!name) return undefined
        text += name
        break
      }
      case 'hour': {
        const hour = hour12 ? ((parts.hour % 12) === 0 ? 12 : parts.hour % 12) : parts.hour
        text += padDateNumber(hour, token.width)
        break
      }
      case 'minute':
        text += padDateNumber(parts.minute, token.width)
        break
      case 'second':
        text += padDateNumber(parts.second, token.width)
        break
      case 'ampm': {
        const pm = parts.hour >= 12
        text += token.form === 'AM/PM' ? (pm ? 'PM' : 'AM') : (pm ? 'P' : 'A')
        break
      }
      case 'literal':
        text += token.text
        break
    }
  }
  if (text.length < 1 || text.length > 32_767 || !isCanonicalDisplayText(text)) return undefined
  return text
}

function padDateNumber(value: number, width: number): string {
  const text = String(value)
  return width >= 2 ? text.padStart(2, '0') : text
}

function parseIsoDateTimeParts(lexical: string): DateTimeParts | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:(?:T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?)|(?:Z|[+-]\d{2}:\d{2}))?$/.exec(lexical)
  if (!match) return undefined
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3])
  const hour = match[4] === undefined ? 0 : Number(match[4])
  const minute = match[5] === undefined ? 0 : Number(match[5])
  const second = match[6] === undefined ? 0 : Number(match[6])
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) return undefined
  if (hour > 23 || minute > 59 || second > 59) return undefined
  const offset = /([+-])(\d{2}):(\d{2})$/.exec(lexical)
  if (offset && (Number(offset[2]) > 24 || Number(offset[3]) > 59 || (Number(offset[2]) === 24 && Number(offset[3]) !== 0))) return undefined
  const weekday = gregorianWeekday(year, month, day)
  if (weekday === undefined) return undefined
  return { year, month, day, hour, minute, second, weekday, dateValid: true }
}

function excelSerialToParts(lexical: string, date1904: boolean, needsDate: boolean): DateTimeParts | undefined {
  const parsed = parseDecimalLexical(lexical)
  if (!parsed || parsed.negative) return undefined
  const divisor = 10n ** BigInt(parsed.scale)
  const whole = parsed.scale === 0 ? parsed.coefficient : parsed.coefficient / divisor
  if (whole > BigInt(Number.MAX_SAFE_INTEGER)) return undefined
  const serial = Number(whole)
  if (!Number.isSafeInteger(serial) || serial < 0) return undefined
  const frac = parsed.scale === 0 ? 0n : parsed.coefficient % divisor
  const roundedSeconds = Number((frac * 86400n + divisor / 2n) / divisor)
  if (!Number.isSafeInteger(roundedSeconds) || roundedSeconds < 0) return undefined
  let extraDays = Math.floor(roundedSeconds / 86400)
  let seconds = roundedSeconds % 86400
  const ymd = excelSerialToYmd(serial + extraDays, date1904)
  const hour = Math.floor(seconds / 3600)
  const minute = Math.floor((seconds % 3600) / 60)
  const second = seconds % 60
  if (!ymd) {
    if (needsDate) return undefined
    return { year: 1899, month: 12, day: 31, hour, minute, second, weekday: 0, dateValid: false }
  }
  const weekday = date1904 ? (serial + extraDays + 5) % 7 : (serial + extraDays - 1) % 7
  if (weekday < 0) return undefined
  return { year: ymd.year, month: ymd.month, day: ymd.day, hour, minute, second, weekday, dateValid: true }
}

function gregorianWeekday(year: number, month: number, day: number): number | undefined {
  if (!Number.isSafeInteger(year) || !Number.isSafeInteger(month) || !Number.isSafeInteger(day)) return undefined
  const offsets = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4]
  const y = month < 3 ? year - 1 : year
  return (y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) + offsets[month - 1]! + day) % 7
}

function excelSerialToYmd(serial: number, date1904: boolean): { year: number; month: number; day: number } | undefined {
  if (!Number.isSafeInteger(serial) || serial < 0) return undefined
  if (date1904) return addDaysToYmd(1904, 1, 1, serial)
  if (serial < 1 || serial === 60) return undefined
  return addDaysToYmd(1899, 12, 31, serial > 60 ? serial - 1 : serial)
}

function addDaysToYmd(year: number, month: number, day: number, days: number): { year: number; month: number; day: number } | undefined {
  let y = year, m = month, d = day, remaining = days
  if (!Number.isSafeInteger(remaining) || remaining < 0) return undefined
  while (remaining > 0) {
    const left = daysInMonth(y, m) - d + 1
    if (remaining < left) { d += remaining; remaining = 0; break }
    remaining -= left
    d = 1
    m += 1
    if (m > 12) { m = 1; y += 1 }
    if (y > 9999) return undefined
  }
  if (y < 1900 || y > 9999) return undefined
  return { year: y, month: m, day: d }
}

function daysInMonth(year: number, month: number): number {
  return [31, year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]!
}

function formatFixedDecimalLexical(lexical: string, fractionDigits: number, decimalShift = 0): string | undefined {
  const parsed = parseDecimalLexical(lexical)
  if (!parsed) return undefined
  let { coefficient, scale } = parsed
  // Percent is an exact decimal shift, not binary floating-point multiplication.
  scale -= decimalShift
  if (scale < fractionDigits) {
    const shift = fractionDigits - scale
    if (shift > 32) return undefined
    coefficient *= 10n ** BigInt(shift)
    scale = fractionDigits
  } else if (scale > fractionDigits) {
    const shift = scale - fractionDigits
    if (shift > 32) return undefined
    const divisor = 10n ** BigInt(shift)
    const remainder = coefficient % divisor
    coefficient = coefficient / divisor
    if (remainder * 2n >= divisor) coefficient += 1n
    scale = fractionDigits
  }
  let digits = coefficient.toString()
  if (fractionDigits > 0 && digits.length <= fractionDigits) digits = digits.padStart(fractionDigits + 1, '0')
  const sign = parsed.negative && coefficient !== 0n ? '-' : ''
  const text = fractionDigits === 0 ? `${sign}${digits}` : `${sign}${digits.slice(0, digits.length - fractionDigits)}.${digits.slice(digits.length - fractionDigits)}`
  if (text.length < 1 || text.length > 32 || !PRINTABLE_ASCII.test(text)) return undefined
  return text
}

function parseDecimalLexical(lexical: string): { negative: boolean; coefficient: bigint; scale: number } | undefined {
  const match = /^([+-])?(?:([0-9]+)(?:\.([0-9]*))?|\.([0-9]+))(?:[eE]([+-]?[0-9]+))?$/.exec(lexical)
  if (!match) return undefined
  const exponent = match[5] === undefined ? 0 : Number(match[5])
  if (!Number.isSafeInteger(exponent) || exponent > 32 || exponent < -32) return undefined
  const intDigits = match[2] ?? ''
  const fracDigits = match[3] ?? match[4] ?? ''
  const raw = `${intDigits}${fracDigits}`.replace(/^0+/, '') || '0'
  if (raw.length > 32) return undefined
  let scale = fracDigits.length - exponent
  let coefficient = BigInt(raw)
  if (scale < 0) {
    const shift = -scale
    if (shift > 32 || raw.length + shift > 32) return undefined
    coefficient *= 10n ** BigInt(shift)
    scale = 0
  }
  return { negative: match[1] === '-', coefficient, scale }
}
