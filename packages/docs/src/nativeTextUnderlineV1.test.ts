import { describe, expect, it } from 'vitest'
import { nativeTextUnderlineCommandsV1 } from './nativeTextUnderlineV1.js'
import type { NativeDocxLineFragmentV1 } from './nativeShapingLines.js'

const fragment = { id: 'fragment:1', source_id: 'run:1', source_kind: 'run', face_id: 'font:1', advance_inline_millipoints: 5000, whitespace: false, underline_position_millipoints: -700, underline_thickness_millipoints: 500 } as NativeDocxLineFragmentV1
describe('font-bound native underline policy', () => {
  it('uses font position and thickness for single/double and inherited text color', () => {
    expect(nativeTextUnderlineCommandsV1('single', '123456', fragment, 'placed:1', 'line:1', 1000, 12000)).toMatchObject({ ok: true, commands: [{ stroke_index: 0, x1_millipoints: 1000, x2_millipoints: 6000, y1_millipoints: 12700, y2_millipoints: 12700, width_millipoints: 500, stroke_rgb: '123456' }] })
    const double = nativeTextUnderlineCommandsV1('double', undefined, fragment, 'placed:1', 'line:1', 1000, 12000)
    expect(double.ok && double.commands.map(command => [command.stroke_index, command.y1_millipoints, command.stroke_rgb])).toEqual([[0,12700,'000000'],[1,13700,'000000']])
  })
  it('supports word-only whitespace omission and explicit tab advances', () => {
    expect(nativeTextUnderlineCommandsV1('words', undefined, { ...fragment, whitespace: true }, 'p','l',1000,12000)).toEqual({ ok: true, commands: [] })
    expect(nativeTextUnderlineCommandsV1('single', undefined, { ...fragment, source_kind: 'tab', whitespace: true }, 'p','l',1000,12000)).toMatchObject({ ok: true, commands: [{ x2_millipoints: 6000 }] })
    expect(nativeTextUnderlineCommandsV1('none', undefined, fragment, 'p','l',1000,12000)).toEqual({ ok: true, commands: [] })
    expect(nativeTextUnderlineCommandsV1('single', undefined, { ...fragment, advance_inline_millipoints: 0 }, 'p','l',1000,12000)).toEqual({ ok: true, commands: [] })
  })
  it('refuses missing font metrics, unknown styles/colors, and overflowing geometry', () => {
    for (const patch of [{ underline_position_millipoints: undefined }, { underline_thickness_millipoints: 0 }, { underline_thickness_millipoints: NaN }, { face_id: undefined }, { source_kind: 'image' as const }]) expect(nativeTextUnderlineCommandsV1('single', undefined, { ...fragment, ...patch }, 'p','l',1000,12000).ok).toBe(false)
    expect(nativeTextUnderlineCommandsV1('wave', undefined, fragment, 'p','l',1000,12000).ok).toBe(false)
    expect(nativeTextUnderlineCommandsV1('single', 'red', fragment, 'p','l',1000,12000).ok).toBe(false)
    for (const x of [-1,Infinity,1e12]) expect(nativeTextUnderlineCommandsV1('single', undefined, fragment, 'p','l',x,12000).ok).toBe(false)
  })
})
