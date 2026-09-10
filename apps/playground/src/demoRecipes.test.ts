import { existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { GuidedRecipe } from './components/GuidedRecipe'
import { DEMOS } from './demoRegistry'
import { DEMO_RECIPES } from './demoRecipes'
import { SURFACES } from './route'

const repositoryRoot = fileURLToPath(new URL('../../../', import.meta.url))
const githubSourceRoot = 'https://github.com/injectinglabs/injoffice/blob/main/'

describe('guided demo recipes', () => {
  it('guides agent tasks through the mock-only controls and expandable evidence', () => {
    const instructions = DEMO_RECIPES.agent.steps.map((step) => step.instruction).join(' ')
    expect(instructions).toContain('Customize this task')
    expect(instructions).toContain('Preview change')
    expect(instructions).toContain('Technical details → Try the safety boundaries')
    expect(instructions).not.toContain('Keep “Built-in mock agent')
  })

  it('covers every proof surface with actionable, ordered steps', () => {
    const registeredSurfaces = SURFACES.filter((surface) => surface !== 'overview').sort()

    expect(Object.keys(DEMO_RECIPES).sort()).toEqual(registeredSurfaces)
    expect(DEMOS.map((demo) => demo.recipe.id)).toHaveLength(16)
    expect(new Set(DEMOS.map((demo) => demo.recipe.id)).size).toBe(16)

    for (const demo of DEMOS) {
      const { recipe } = demo
      expect(recipe).toBe(DEMO_RECIPES[demo.surface])
      expect(recipe.outcome.length).toBeGreaterThan(40)
      expect(recipe.minutes).toBeGreaterThanOrEqual(1)
      expect(recipe.steps.length).toBeGreaterThanOrEqual(3)
      expect(recipe.steps.length).toBeLessThanOrEqual(5)
      expect(new Set(recipe.steps.map((step) => step.id)).size).toBe(recipe.steps.length)
      expect(recipe.steps.every((step) => step.title.length > 5 && step.instruction.length > 30 && step.evidence.length > 30)).toBe(true)
    }
  })

  it('links to exact repository files and only registered related surfaces', () => {
    const demoSurfaces = new Set(DEMOS.map((demo) => demo.surface))

    for (const demo of DEMOS) {
      expect(demo.recipe.sources.length).toBeGreaterThanOrEqual(2)
      expect(demo.recipe.related.length).toBeGreaterThanOrEqual(1)
      expect(new Set(demo.recipe.related.map((relation) => relation.surface)).size).toBe(demo.recipe.related.length)

      for (const item of demo.recipe.sources) {
        expect(item.href).toBe(`${githubSourceRoot}${item.path}`)
        expect(existsSync(`${repositoryRoot}${item.path}`), `${item.path} should be an exact repository path`).toBe(true)
      }

      for (const relation of demo.recipe.related) {
        expect(demoSurfaces.has(relation.surface)).toBe(true)
        expect(relation.surface).not.toBe(demo.surface)
        expect(relation.reason.length).toBeGreaterThan(30)
      }
    }
  })

  it('renders an accessible first step, progress, sources, and continuation links', () => {
    const demo = DEMOS.find((item) => item.surface === 'sheets')!
    const markup = renderToStaticMarkup(createElement(GuidedRecipe, { recipe: demo.recipe, accent: demo.accent }))

    expect(markup).toContain('guided-recipe--mint')
    expect(markup).toContain('role="progressbar"')
    expect(markup).toContain('aria-valuenow="0"')
    expect(markup).toContain('aria-current="step"')
    expect(markup).toContain(demo.recipe.steps[0].instruction)
    expect(markup).toContain(demo.recipe.sources[0].href)
    expect(markup).toContain('href="#/formulas"')
    expect(markup).toContain('Mark done &amp; continue')
  })
})
