export type NativeCompactGeneralPreviewV1 =
  | { readonly status: 'ready'; readonly text: string; readonly policy: 'host-general-seven-significant-v1' }
  | { readonly status: 'refused'; readonly warning: string }

/** Explicit host display policy, NOT Excel General or a source number format.
 * Decimal rounding is half away from zero, at seven significant digits.
 * Scientific notation is chosen below 1e-6 or at/above 1e7. No locale, binary
 * floating-point conversion, formula evaluation, or source mutation occurs.
 */
export function compactNativeGeneralNumberPreviewV1(lexical: string): NativeCompactGeneralPreviewV1 {
  const refuse = (): NativeCompactGeneralPreviewV1 => ({ status: 'refused', warning: 'Compact number preview unavailable; showing the stored value.' })
  if (typeof lexical !== 'string' || lexical.length < 1 || lexical.length > 128) return refuse()
  const match = /^([+-])?(?:([0-9]+)(?:\.([0-9]*))?|\.([0-9]+))(?:[eE]([+-]?)([0-9]{1,3}))?$/.exec(lexical)
  if (!match) return refuse()
  let exponent = 0
  for (const digit of match[6] ?? '') exponent = exponent * 10 + digit.charCodeAt(0) - 48
  if (exponent > 128) return refuse()
  if (match[5] === '-') exponent = -exponent
  const whole = match[2] ?? '', raw = whole + (match[3] ?? match[4] ?? '')
  const leading = raw.length - raw.replace(/^0+/, '').length
  let digits = raw.slice(leading), point = whole.length + exponent - leading
  if (!digits) return { status: 'ready', text: '0', policy: 'host-general-seven-significant-v1' }
  if (digits.length > 7) {
    const roundUp = digits[7]! >= '5'
    digits = digits.slice(0, 7)
    if (roundUp) digits = (BigInt(digits) + 1n).toString()
    if (digits.length > 7) { point++; digits = digits.slice(0, 7) }
  }
  digits = digits.replace(/0+$/, '')
  let text: string
  if (point > 7 || point <= -6) {
    const exp = point - 1
    text = digits[0]! + (digits.length > 1 ? '.' + digits.slice(1) : '') + 'e' + (exp >= 0 ? '+' : '') + exp
  } else if (point <= 0) text = '0.' + '0'.repeat(-point) + digits
  else if (point >= digits.length) text = digits + '0'.repeat(point - digits.length)
  else text = digits.slice(0, point) + '.' + digits.slice(point)
  return { status: 'ready', text: (match[1] === '-' ? '-' : '') + text, policy: 'host-general-seven-significant-v1' }
}
