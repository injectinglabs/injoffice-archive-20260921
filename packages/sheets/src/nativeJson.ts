import { XLSX_NATIVE_RESOURCE_LIMITS } from './nativeContract.generated.js'

/** Parse JSON without losing duplicate-key evidence. This module intentionally
 * depends only on ECMAScript primitives so it is safe in workers and servers. */
export function parseNativeWorkbookJson(source: string): unknown {
  if (typeof source !== 'string') throw new TypeError('native XLSX JSON must be a string')
  const bytes = utf8ByteLength(source)
  if (bytes < 1 || bytes > XLSX_NATIVE_RESOURCE_LIMITS.maxJsonBytes) {
    throw new SyntaxError(`native XLSX JSON size must be 1..${XLSX_NATIVE_RESOURCE_LIMITS.maxJsonBytes} UTF-8 bytes`)
  }
  const scanner = new StrictJsonScanner(source)
  scanner.scan()
  return JSON.parse(source) as unknown
}

export function utf8ByteLength(value: string): number {
  let bytes = 0
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index)
    if (unit <= 0x7f) bytes++
    else if (unit <= 0x7ff) bytes += 2
    else if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = value.charCodeAt(++index)
      if (!(low >= 0xdc00 && low <= 0xdfff)) throw new SyntaxError('native XLSX JSON contains an unpaired UTF-16 surrogate')
      bytes += 4
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      throw new SyntaxError('native XLSX JSON contains an unpaired UTF-16 surrogate')
    } else bytes += 3
  }
  return bytes
}

class StrictJsonScanner {
  private offset = 0
  private tokens = 0
  private readonly number = /-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?/y

  constructor(private readonly source: string) {}

  scan(): void {
    this.whitespace()
    this.value(0, '')
    this.whitespace()
    if (this.offset !== this.source.length) throw this.error('trailing JSON value')
  }

  private value(depth: number, path: string): void {
    if (depth > XLSX_NATIVE_RESOURCE_LIMITS.maxJsonDepth) throw this.error(`JSON nesting exceeds ${XLSX_NATIVE_RESOURCE_LIMITS.maxJsonDepth}`)
    this.claimToken()
    const character = this.source[this.offset]
    if (character === '{') return this.object(depth, path)
    if (character === '[') return this.array(depth, path)
    if (character === '"') { this.string(); return }
    if (character === 't') return this.literal('true')
    if (character === 'f') return this.literal('false')
    if (character === 'n') return this.literal('null')
    this.number.lastIndex = this.offset
    const match = this.number.exec(this.source)
    if (!match) throw this.error('invalid JSON value')
    this.offset += match[0].length
  }

  private object(depth: number, path: string): void {
    this.offset++
    this.whitespace()
    const keys = new Set<string>()
    if (this.source[this.offset] === '}') { this.offset++; this.claimToken(); return }
    while (true) {
      if (this.source[this.offset] !== '"') throw this.error('object key must be a string')
      this.claimToken()
      const key = this.string()
      if (keys.has(key)) throw this.error(`duplicate object key ${JSON.stringify(key)} at ${path || '/'}`)
      keys.add(key)
      this.whitespace()
      if (this.source[this.offset++] !== ':') throw this.error('object key must be followed by a colon')
      this.whitespace()
      this.value(depth + 1, `${path}/${escapePointer(key)}`)
      this.whitespace()
      const delimiter = this.source[this.offset++]
      if (delimiter === '}') { this.claimToken(); return }
      if (delimiter !== ',') throw this.error('object entries must be separated by a comma')
      this.whitespace()
    }
  }

  private array(depth: number, path: string): void {
    this.offset++
    this.whitespace()
    if (this.source[this.offset] === ']') { this.offset++; this.claimToken(); return }
    let index = 0
    while (true) {
      this.value(depth + 1, `${path}/${index++}`)
      this.whitespace()
      const delimiter = this.source[this.offset++]
      if (delimiter === ']') { this.claimToken(); return }
      if (delimiter !== ',') throw this.error('array entries must be separated by a comma')
      this.whitespace()
    }
  }

  private string(): string {
    const start = this.offset
    this.offset++
    while (this.offset < this.source.length) {
      const code = this.source.charCodeAt(this.offset++)
      if (code === 0x22) {
        const raw = this.source.slice(start, this.offset)
        return JSON.parse(raw) as string
      }
      if (code < 0x20) throw this.error('unescaped control character in JSON string')
      if (code !== 0x5c) continue
      const escape = this.source[this.offset++]
      if (escape === 'u') {
        if (!/^[0-9a-fA-F]{4}$/.test(this.source.slice(this.offset, this.offset + 4))) throw this.error('invalid JSON Unicode escape')
        this.offset += 4
      } else if (!escape || !'"\\/bfnrt'.includes(escape)) throw this.error('invalid JSON escape')
    }
    throw this.error('unterminated JSON string')
  }

  private literal(value: string): void {
    if (this.source.slice(this.offset, this.offset + value.length) !== value) throw this.error(`invalid JSON literal`)
    this.offset += value.length
  }

  private whitespace(): void {
    while (this.offset < this.source.length && /[\u0009\u000a\u000d\u0020]/.test(this.source[this.offset])) this.offset++
  }

  private claimToken(): void {
    this.tokens++
    if (this.tokens > XLSX_NATIVE_RESOURCE_LIMITS.maxJsonTokens) throw this.error(`JSON token count exceeds ${XLSX_NATIVE_RESOURCE_LIMITS.maxJsonTokens}`)
  }

  private error(message: string): SyntaxError { return new SyntaxError(`${message} at byte ${this.offset}`) }
}

function escapePointer(value: string): string { return value.replace(/~/g, '~0').replace(/\//g, '~1') }
