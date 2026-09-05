import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const editor = readFileSync(new URL('./UniverEditor.tsx', import.meta.url), 'utf8')

describe('playground Univer editor lifecycle', () => {
  it('does not construct Univer inside the StrictMode effect probe', () => {
    expect(editor).toContain('useEffect(() => deferNestedReactRootStart(() => {')
    expect(editor).toContain('if (!containerRef.current) return')
  })

  it('stops integrations before disposing the nested Univer root after the commit', () => {
    const managerStop = editor.indexOf('manager.stop()')
    const univerDispose = editor.indexOf('window.setTimeout(() => univer.dispose(), 0)')
    expect(managerStop).toBeGreaterThan(-1)
    expect(univerDispose).toBeGreaterThan(managerStop)
  })

  it('hydrates chart float DOMs only after Univer mounts its render engine', () => {
    const renderedHook = editor.indexOf('getHooks().onRendered')
    const chartHydration = editor.indexOf('requestAnimationFrame(mountFileCharts)', renderedHook)
    expect(renderedHook).toBeGreaterThan(-1)
    expect(chartHydration).toBeGreaterThan(renderedHook)
    expect(editor).toContain('renderedHook.dispose()')
  })
})
