import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const source = dirname(fileURLToPath(import.meta.url))
const css = readFileSync(resolve(source, 'workbench.css'), 'utf8')
const main = readFileSync(resolve(source, 'main.tsx'), 'utf8')
const app = readFileSync(resolve(source, 'App.tsx'), 'utf8')

describe('playground workbench design system', () => {
  it('loads one app-level override after the legacy proof surface styles', () => {
    expect(main.indexOf("'./workbench.css'")).toBeGreaterThan(main.indexOf("'./studio.css'"))
    expect(existsSync(resolve(source, 'panel.css'))).toBe(false)
  })

  it('defines a shared control, state, typography, and workspace vocabulary', () => {
    for (const token of [
      '--workspace',
      '--paper',
      '--action',
      '--live',
      '--warning',
      '--danger',
      '--control-height',
      '--radius-control',
      '--type-body',
      '--type-mono',
    ]) expect(css).toContain(token)

    expect(css).toContain('.tool-page__controls > button')
    expect(css).toContain('.native-toolbar button')
    expect(css).toContain('.capability-editor-actions button')
    expect(css).toContain('.view-switcher button')
    expect(css).toContain('.ioc-panel button:not(.ioc-chip)')
    expect(css).toContain('.ios-deck-canvas button')
    expect(css).toContain('.ioc-outline-gutter button')
    expect(css).toContain('.tool-page__controls')
    expect(css).toContain('.univer-toolbar__group')
    expect(css).toContain('.pdf-page-input__control')
    expect(css).toContain('.collab-sim-slides-tools')
    expect(css).toContain('.collab-sim-pdf-tools')
    expect(css).toContain('.formula-example')
    expect(css).toContain('align-items: flex-end')
    expect(css).toContain(':focus-visible')
    expect(css).toContain('prefers-reduced-motion')
    expect(css).toContain('html[data-theme="dark"]')
    expect(css).toContain('color-scheme: dark')
    expect(css).toContain('--chrome')
    expect(css).toContain('.scheme-toggle')
  })

  it('exposes the active tool and accent as stable styling hooks', () => {
    expect(app).toContain('data-surface={surface}')
    expect(app).toContain('data-workbench-surface={surface}')
    expect(app).toContain('data-accent={demo.accent}')
    expect(app).toContain('ColorSchemeToggle')
    expect(app).toContain('persistColorScheme')
  })

  it('uses a continuous showcase layout with independently identified live sections', () => {
    expect(app).toContain('data-layout="univer"')
    expect(app).toContain('className="app-frame"')
    expect(app).toContain('data-navigation="scroll"')
    expect(app).toContain('className="demo-section"')
    expect(app).toContain('id={`demo-${section.key}`}')
    expect(app).toContain('aria-labelledby={`demo-title-${section.key}`}')
    expect(app).toContain('data-scroll-section={section.key}')
    expect(app).toContain('{demo.navTitle}')
    expect(app).toContain('AGENT_TOOLS')
    expect(app).toContain('agentHref')
    expect(app).toContain('AI change sets, ${item.label}')
    expect(app).toContain('agentTool?.title ?? demo.title')
    expect(app).toContain('className="demo-breadcrumb" aria-label="Breadcrumb"')
    expect(app).toContain('className="demo-context-actions"')
    expect(app).toContain('className="demo-preview"')
    expect(app).toContain('SCROLL_SECTIONS.filter(section => section.demo).map(section => <DemoSection')
    expect(app).toContain('activeSectionKey')
    expect(app).toContain("aria-current={surface === demo.surface ? 'location' : undefined}")
    expect(app).not.toContain('navigationCollapsed')
    expect(app).not.toContain('function NavigationToggle')
    expect(css).toContain('.app-frame')
    expect(css).toContain('.demo-preview')
    expect(css).not.toContain('.app-shell[data-navigation="collapsed"]')
  })

  it('provides an accessible source and proof drawer from demo metadata', () => {
    expect(app).toContain('function SourceProofDrawer')
    expect(app).toContain('<GuidedRecipe recipe={demo.recipe}')
    expect(app).toContain('<DemoSource source={demo.recipe.sources[0]}')
    expect(app).toContain('aria-haspopup="dialog"')
    expect(app).toContain('role="dialog"')
    expect(app).toContain('aria-modal="true"')
    expect(app).toContain("event.key === 'Escape'")
    expect(app).toContain("event.key !== 'Tab'")
    expect(app).toContain('returnFocusRef.current?.focus()')
    expect(css).toContain('.source-proof-drawer')
    expect(css).toMatch(/@media \(prefers-reduced-motion: reduce\)[\s\S]*\.source-proof-drawer \{ animation: none; \}/)
  })

  it('drives Univer dark mode from the playground color scheme', () => {
    const editor = readFileSync(resolve(source, 'UniverEditor.tsx'), 'utf8')
    const simulator = readFileSync(resolve(source, 'collabSimulator.tsx'), 'utf8')
    expect(editor).toContain('darkMode: univerDarkMode()')
    expect(editor).toContain('bindUniverColorScheme(univerAPI)')
    expect(simulator).toContain('darkMode: univerDarkMode()')
    expect(simulator).toContain('bindUniverColorScheme(univerAPI)')
  })
})
