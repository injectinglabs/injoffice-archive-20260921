import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createPlaygroundHistory, playgroundHistoryAuthors } from './historyLifecycle'
import { createPlaygroundHistoryController } from './historyTimeline'

describe('playground history timeline wiring', () => {
  it('captures live content and activates restore through the timeline controller', async () => {
    const { manager } = createPlaygroundHistory([
      { snapshot: 'one', contentType: 'text/plain', description: 'first' },
    ])
    const authors = playgroundHistoryAuthors()
    let live = 'two'
    let preview = ''
    let activated = ''
    const controller = createPlaygroundHistoryController(
      manager,
      () => ({ snapshot: live, contentType: 'text/plain' }),
      (loaded) => { preview = loaded.snapshot },
      (snapshot) => { activated = snapshot },
    )

    const captured = await controller.capture({ author: authors.user, description: 'from timeline' })
    expect(captured.description).toBe('from timeline')
    expect((await manager.loadPreview(captured.id)).snapshot).toBe('two')

    await controller.preview('v1')
    expect(preview).toBe('one')

    const restored = await controller.restore({ sourceVersionId: 'v1', author: authors.user })
    expect(restored.sourceVersionId).toBe('v1')
    expect(activated).toBe('one')
  })

  it('mounts the package timeline on the history demo page', () => {
    const source = readFileSync(new URL('./pages/HistoryPage.tsx', import.meta.url), 'utf8')
    expect(source).toContain('HistoryTimeline')
    expect(source).toContain('createPlaygroundHistoryController')
  })
})

