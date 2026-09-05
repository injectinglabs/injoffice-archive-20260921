import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const pageSource = (name: string) => readFileSync(new URL(`./pages/${name}.tsx`, import.meta.url), 'utf8')
const packageSource = (path: string) => readFileSync(new URL(`../../../packages/${path}`, import.meta.url), 'utf8')

describe('focused demo trust contracts', () => {
  it('renders every current chart type through a lifecycle-owned ECharts surface', () => {
    const page = pageSource('ChartsPage')
    const preview = packageSource('charts/src/EChartsPreview.tsx')

    expect(page).toContain('<EChartsPreview')
    expect(page).toContain('option={option}')
    expect(page).toContain('{CHART_TYPES.length} chart types')
    expect(page).toContain('START_LOW')
    expect(page).toContain('START_HIGH')
    expect(page).toContain('LINK_CATEGORIES')
    expect(page).toContain('HIERARCHY_CATEGORIES')
    expect(page).not.toContain('chart-proof__bars')
    expect(preview).toContain("import * as echarts from 'echarts'")
    expect(preview).toContain('new ResizeObserver')
    expect(preview).toContain('new MutationObserver')
    expect(preview).toContain('setOption(option, { notMerge: true })')
    expect(preview).toContain('instanceRef.current?.dispose()')
    expect(preview).toContain('role="img"')
  })

  it('uses the package shape dispatcher and exposes simplified-preview labels', () => {
    const page = pageSource('ShapesPage')

    expect(page).toContain('<ShapePreview')
    expect(page).toContain('shapePreviewFidelity')
    expect(page).toContain('Approximated browser preview')
    expect(page).toContain('Generic browser preview')
    expect(page).not.toContain('function geometry(')
    expect(page).not.toContain('roundRectPath')
  })
})
