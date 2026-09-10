import { describe, expect, it } from 'vitest'
import { nativeTextHighlightCommandV1 } from './nativeTextHighlightV1.js'
import type { NativeDocxLineFragmentV1 } from './nativeShapingLines.js'

const fragment = { id: 'fragment:1', source_id: 'run:1', source_kind: 'run', advance_inline_millipoints: 5000, ascent_millipoints: 8000, descent_millipoints: -2000 } as NativeDocxLineFragmentV1
describe('native text highlights', () => {
  it('maps authored colors and exact shaped metric geometry', () => {
    const expected = { black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00', magenta: 'FF00FF', red: 'FF0000', yellow: 'FFFF00', white: 'FFFFFF', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0' }
    for (const [color, fill] of Object.entries(expected)) expect(nativeTextHighlightCommandV1(color,fragment,'placed:1','line:1',1000,12000)).toMatchObject({ ok: true, command: { fill_rgb: fill, x_millipoints: 1000, y_millipoints: 4000, width_millipoints: 5000, height_millipoints: 10000 } })
  })
  it('omits none and zero-advance backgrounds without fabricating width', () => {
    expect(nativeTextHighlightCommandV1('none',fragment,'placed:1','line:1',1000,12000)).toEqual({ ok: true })
    expect(nativeTextHighlightCommandV1('yellow',{ ...fragment, advance_inline_millipoints: 0 },'placed:1','line:1',1000,12000)).toEqual({ ok: true })
    expect(nativeTextHighlightCommandV1('yellow',{ ...fragment, text: ' ', whitespace: true },'placed:1','line:1',1000,12000)).toMatchObject({ ok: true, command: { width_millipoints: 5000 } })
  })
  it('refuses unknown colors, non-run highlights and unbounded geometry', () => {
    expect(nativeTextHighlightCommandV1('constructor',fragment,'placed:1','line:1',1000,12000).ok).toBe(false)
    for (const source_kind of ['tab', 'list-marker', 'image'] as const) expect(nativeTextHighlightCommandV1('yellow',{ ...fragment, source_kind },'placed:1','line:1',1000,12000).ok).toBe(false)
    for (const x of [-1,Infinity,1e12]) expect(nativeTextHighlightCommandV1('yellow',fragment,'placed:1','line:1',x,12000).ok).toBe(false)
    expect(nativeTextHighlightCommandV1('yellow',fragment,'placed:1','line:1',1000,100).ok).toBe(false)
  })
})
