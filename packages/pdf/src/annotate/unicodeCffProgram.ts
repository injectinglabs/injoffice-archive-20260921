/** Bounded Type 2 preflight for programs preserved in a CID CFF wrapper.
 * In particular, name-keyed endchar composites cannot be copied into CID fonts.
 * Adobe Technical Note 5177 defines the stack, subroutine and hint-mask rules. */
export function validateCffPrograms(glyphs: readonly Uint8Array[], globals: readonly Uint8Array[], locals: readonly (readonly Uint8Array[])[], selected: readonly number[]): void {
  let budget = 5_000_000
  const bias = (count: number) => count < 1240 ? 107 : count < 33900 ? 1131 : 32768
  for (let gid = 0; gid < glyphs.length; gid++) {
    const stack: number[] = [], transient = new Array<number | undefined>(32)
    let hints = 0, widthSeen = false, moved = false, masked = false, hintMasked = false, verticalHints = false, ended = false
    const pop = () => { const n = stack.pop(); if (n === undefined) throw new Error('CFF operand stack underflow'); return n }
    const integer = (n: number) => { if (!Number.isInteger(n)) throw new Error('invalid CFF integer operand'); return n }
    const width = (required: number) => {
      if (!widthSeen && stack.length > required) stack.shift()
      widthSeen = true
    }
    const stems = () => {
      if (!widthSeen && stack.length % 2) stack.shift()
      widthSeen = true
      if (stack.length % 2) throw new Error('invalid CFF stem operands')
      if (stack.some((n, i) => i % 2 === 1 && n < 0 && n !== -20 && n !== -21)) throw new Error('undefined negative CFF stem width')
      hints += stack.length / 2; stack.length = 0
      if (hints > 96) throw new Error('CFF hint limit exceeded')
    }
    const parse = (bytes: Uint8Array, depth: number): void => {
      if (depth > 10) throw new Error('CFF subroutine depth exceeded')
      let offset = 0
      const next = () => { if (offset >= bytes.length) throw new Error('truncated CFF charstring'); return bytes[offset++]! }
      while (offset < bytes.length && !ended) {
        if (--budget < 0) throw new Error('CFF program operation limit exceeded')
        const op = next()
        if (op >= 32 || op === 28) {
          let n: number
          if (op === 28) { n = next() * 256 + next(); if (n & 0x8000) n -= 0x10000 }
          else if (op <= 246) n = op - 139
          else if (op <= 250) n = (op - 247) * 256 + next() + 108
          else if (op <= 254) n = -(op - 251) * 256 - next() - 108
          else { const a = next() * 0x1000000 + next() * 0x10000 + next() * 256 + next(); n = (a | 0) / 65536 }
          stack.push(n)
        } else if (op === 10 || op === 29) {
          const subrs = op === 29 ? globals : locals[selected[gid]!]!
          const index = integer(pop()) + bias(subrs.length)
          if (!subrs[index]) throw new Error('invalid CFF subroutine index')
          parse(subrs[index]!, depth + 1)
        } else if (op === 11) {
          if (!depth || offset !== bytes.length) throw new Error('invalid CFF subroutine return')
          return
        } else if ([1, 3, 18, 23, 19, 20].includes(op)) {
          const isMask = op === 19 || op === 20
          if ((!isMask && (moved || masked)) || op === 20 && (moved || hintMasked) || isMask && (moved || masked) && stack.length > 0 || [1, 18].includes(op) && verticalHints) throw new Error('invalid CFF hint phase')
          const previousHints = hints
          stems()
          if (!isMask && hints === previousHints) throw new Error('CFF stem operator has no operands')
          if ([3, 23].includes(op) || isMask && hints > previousHints) verticalHints = true
          if (isMask) {
            if (!hints) throw new Error('CFF hint mask has no stem hints')
            for (let i = 0; i < Math.ceil(hints / 8); i++) next()
            masked = true
            if (op === 19) hintMasked = true
          }
        } else if (op === 14) {
          if (stack.length > 1) throw new Error('name-keyed CFF endchar composites require conversion before CID embedding')
          if (stack.length && widthSeen) throw new Error('invalid CFF endchar operands')
          stack.length = 0; ended = true
          if (offset !== bytes.length) throw new Error('trailing CFF charstring bytes')
        } else if ([4, 21, 22].includes(op)) {
          const count = op === 21 ? 2 : 1
          width(count)
          if (stack.length !== count) throw new Error('invalid CFF moveto operands')
          stack.length = 0; moved = true
        } else if ([5, 6, 7, 8, 24, 25, 26, 27, 30, 31].includes(op)) {
          const n = stack.length
          const valid = op === 5 ? n >= 2 && n % 2 === 0 : op === 6 || op === 7 ? n >= 1 : op === 8 ? n >= 6 && n % 6 === 0 : op === 24 ? n >= 8 && (n - 2) % 6 === 0 : op === 25 ? n >= 8 && (n - 6) % 2 === 0 : n >= 4 && (n % 4 === 0 || n % 4 === 1)
          if (!moved || !valid) throw new Error('invalid CFF drawing operands or missing moveto')
          stack.length = 0
        } else if (op === 12) {
          const escaped = next()
          if ([34, 35, 36, 37].includes(escaped)) {
            if (!moved || stack.length !== ({ 34: 7, 35: 13, 36: 9, 37: 11 } as Record<number, number>)[escaped]) throw new Error('invalid CFF flex operands or missing moveto')
            stack.length = 0
          } else if ([3, 4, 10, 11, 12, 15, 24].includes(escaped)) {
            const b = pop(), a = pop()
            stack.push(escaped === 3 ? Number(Boolean(a) && Boolean(b)) : escaped === 4 ? Number(Boolean(a) || Boolean(b)) : escaped === 10 ? a + b : escaped === 11 ? a - b : escaped === 12 ? a / b : escaped === 15 ? Number(a === b) : a * b)
          } else if ([5, 9, 14, 26].includes(escaped)) {
            const a = pop(); stack.push(escaped === 5 ? Number(!a) : escaped === 9 ? Math.abs(a) : escaped === 14 ? -a : Math.sqrt(a))
          } else if (escaped === 18) pop()
          else if (escaped === 20) { const i = integer(pop()), value = pop(); if (i < 0 || i >= 32) throw new Error('invalid CFF transient index'); transient[i] = value }
          else if (escaped === 21) { const i = integer(pop()); if (i < 0 || i >= 32 || transient[i] === undefined) throw new Error('invalid or uninitialized CFF transient index'); stack.push(transient[i]!) }
          else if (escaped === 22) { const v2 = pop(), v1 = pop(), s2 = pop(), s1 = pop(); stack.push(v1 <= v2 ? s1 : s2) }
          else if (escaped === 27) { const value = pop(); stack.push(value, value) }
          else if (escaped === 28) { const b = pop(), a = pop(); stack.push(b, a) }
          else if (escaped === 29) { const i = Math.max(0, integer(pop())); const value = stack[stack.length - i - 1]; if (value === undefined) throw new Error('CFF index exceeds operand stack'); stack.push(value) }
          else if (escaped === 30) { const j = integer(pop()), n = integer(pop()); if (n < 0 || n > stack.length) throw new Error('invalid CFF roll operands'); if (n) { const values = stack.splice(-n), shift = ((j % n) + n) % n; stack.push(...values.slice(n - shift), ...values.slice(0, n - shift)) } }
          else throw new Error('unsupported CFF computational operator')
        } else throw new Error('unsupported CFF charstring operator')
        if (stack.length > 48 || stack.some(n => !Number.isFinite(n) || Math.abs(n) > 0x7fffffff)) throw new Error('invalid CFF operand stack')
      }
      if (!ended) throw new Error('unterminated CFF charstring or subroutine')
    }
    parse(glyphs[gid]!, 0)
    if (!ended) throw new Error('unterminated CFF glyph')
  }
}
