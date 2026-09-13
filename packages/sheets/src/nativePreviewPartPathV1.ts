/** Validate an exact package part identity, never a URL or filesystem path.
 * Unicode, spaces and percent signs are preserved; no decoding or normalization.
 */
export function isNativePreviewPartPathV1(value: unknown, allowEmpty = false): value is string {
  if (typeof value !== 'string' || value.length > 1024 || /[\u0000-\u001f\u007f]/.test(value)) return false
  if (allowEmpty && value === '') return true
  return value.length > 0 && !value.startsWith('/') && !value.includes('\\') &&
    value.split('/').every(segment => segment !== '' && segment !== '.' && segment !== '..')
}
