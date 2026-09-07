import { describe, expect, it } from 'vitest'
import { DEMO_FORMATS, DEMOS, DEMO_TASKS } from './demoRegistry'

describe('playground capability registry', () => {
  it('registers fifteen focused proof surfaces without implying one page per package', () => {
    expect(DEMOS).toHaveLength(15)
    expect(new Set(DEMOS.map((demo) => demo.surface)).size).toBe(15)
    expect(new Set(DEMOS.map((demo) => demo.packageName)).size).toBe(15)
    expect(DEMOS.every((demo) => demo.description.length > 40)).toBe(true)
    expect(DEMOS.every((demo) => demo.tasks.length > 0 && demo.formats.length > 0)).toBe(true)
    expect(new Set(DEMOS.flatMap((demo) => demo.tasks))).toEqual(new Set(DEMO_TASKS))
    expect(new Set(DEMOS.flatMap((demo) => demo.formats))).toEqual(new Set(DEMO_FORMATS))
  })

  it('keeps interactive document controls distinct from package-only capabilities', () => {
    const docs = DEMOS.find((demo) => demo.surface === 'docs')!
    const sheets = DEMOS.find((demo) => demo.surface === 'sheets')!
    const pptx = DEMOS.find((demo) => demo.surface === 'pptx-native')!
    const pdf = DEMOS.find((demo) => demo.surface === 'pdf')!
    const collab = DEMOS.find((demo) => demo.surface === 'collab')!
    const typography = DEMOS.find((demo) => demo.surface === 'font-metrics')!
    const charts = DEMOS.find((demo) => demo.surface === 'charts')!
    const shapes = DEMOS.find((demo) => demo.surface === 'shapes')!

    expect(docs.description).toMatch(/guarded DOCX text run/i)
    expect(docs.description).toMatch(/verify preservation evidence/i)
    expect(docs.runtime).toBe('Browser')
    expect(pdf.description).toMatch(/page operations/i)
    expect(pdf.description).toMatch(/PDFium/i)
    expect(collab.description).toMatch(/two independent Univer sheets/i)
    expect(collab.description).toMatch(/deck/i)
    expect(collab.description).toMatch(/document/i)
    expect(collab.description).toMatch(/PDF/)
    expect(collab.runtime).toBe('Browser')
    expect(sheets.runtime).toBe('Browser')
    expect(pptx.runtime).toBe('Browser')
    expect(pptx.description).toMatch(/browser-local PPTX bytes/i)
    expect(pptx.description).toMatch(/four exact AutoShapes/i)
    expect(pptx.description).toMatch(/exact output/i)
    expect(typography.runtime).toBe('Browser')
    expect(typography.description).toMatch(/Node boundary/i)
    expect(charts.description).toMatch(/real ECharts adapter/i)
    expect(shapes.description).toMatch(/approximate, and generic previews labeled/i)
  })
})
