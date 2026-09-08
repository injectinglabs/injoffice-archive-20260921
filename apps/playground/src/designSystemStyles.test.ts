import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { DsButton } from './design-system/primitives'

const source = dirname(fileURLToPath(import.meta.url))
const read = (file: string) => readFileSync(resolve(source, file), 'utf8')

describe('playground design-system layering', () => {
  it('keeps application colors in workbench.css and aliases component tokens', () => {
    const tokens = read('design-system/tokens.css')

    expect(tokens).toContain('--ds-paper: var(--paper, #ffffff)')
    expect(tokens).toContain('--ds-select: var(--action, #0f6f8f)')
    expect(tokens).toContain('--ds-type: var(--type-body')
    expect(tokens).not.toContain('html[data-theme="dark"] .ds')
    expect(read('studio.css')).not.toMatch(/(^|\n):root\s*\{/)
  })

  it('loads base surface styles once and leaves route sheets extension-only', () => {
    const components = read('design-system/components.css')
    const liveCreateEdit = read('design-system/live-create-edit.css')
    const liveTools = read('design-system/live-tools.css')

    expect(components).toContain("@import './surfaces.css';")
    expect(components.match(/\.ds-btn\s*\{/g)).toHaveLength(1)
    for (const routeCss of [liveCreateEdit, liveTools]) {
      expect(routeCss).not.toMatch(/@import ['"].*?(tokens|components|surfaces)\.css['"]/)
      expect(routeCss).not.toMatch(/\.ds-btn(?:[.:\[\s,{])/)
    }
  })

  it('gives design-system buttons one class authority while retaining extensions', () => {
    const markup = renderToStaticMarkup(createElement(DsButton, {
      variant: 'filled',
      className: 'workbench-button workbench-button--primary analytics-trigger',
      children: 'Run',
    }))

    expect(markup).toContain('class="ds-btn ds-btn--filled analytics-trigger"')
    expect(markup).not.toContain('workbench-button')
    expect(read('workbench.css')).toContain('button:not(.ds-btn)')
  })

  it('reserves a viewport-aware workspace and lets reading surfaces rejoin document flow', () => {
    const workbench = read('workbench.css')

    expect(workbench).toContain('--workspace-reserved-min: clamp(480px, 72dvh, 640px)')
    expect(workbench).toContain('--workspace-reserved-min: clamp(360px, 60dvh, 520px)')
    expect(workbench).toContain('height: clamp(var(--workspace-reserved-min), calc(100dvh - 160px), 880px)')
    expect(workbench).toMatch(/\.demo-section \.demo-stage \{[\s\S]*?overflow: hidden;/)
    expect(workbench).toContain('scrollbar-gutter: stable')
    expect(workbench).toContain('[data-scroll-section^="agent-"][data-section-loaded="true"] .demo-preview')
    expect(workbench).toContain('[data-scroll-section="docs"]')
    expect(workbench).toContain('[data-scroll-section="pdf"]')
  })
})
