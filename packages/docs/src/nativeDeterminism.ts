/** Locale-independent ordering for wire-visible native diagnostics and issues. */
export function compareNativeCodeUnits(left: string, right: string): number {
  if (left === right) return 0
  const length = Math.min(left.length, right.length)
  for (let index = 0; index < length; index += 1) {
    const difference = left.charCodeAt(index) - right.charCodeAt(index)
    if (difference !== 0) return difference
  }
  return left.length - right.length
}

/** Host-Unicode-independent ASCII folding; every non-ASCII UTF-16 code unit is preserved. */
export function asciiLowerNative(value: string): string {
  let output = ''
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index)
    output += String.fromCharCode(codeUnit >= 0x41 && codeUnit <= 0x5a ? codeUnit + 0x20 : codeUnit)
  }
  return output
}

export function asciiUpperNative(value: string): string {
  let output = ''
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index)
    output += String.fromCharCode(codeUnit >= 0x61 && codeUnit <= 0x7a ? codeUnit - 0x20 : codeUnit)
  }
  return output
}

export function compareNativeValidationIssues(
  left: { path: string; code: string; message: string },
  right: { path: string; code: string; message: string },
): number {
  return compareNativeCodeUnits(left.path, right.path)
    || compareNativeCodeUnits(left.code, right.code)
    || compareNativeCodeUnits(left.message, right.message)
}
