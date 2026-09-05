import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..')

describe('Injecting logo', () => {
  it('ships the same mark for GitHub and the playground', () => {
    const github = readFileSync(resolve(root, 'logo.svg'), 'utf8')
    const playground = readFileSync(resolve(root, 'apps/playground/public/logo.svg'), 'utf8')
    expect(existsSync(resolve(root, 'logo.png'))).toBe(true)
    expect(github).toContain('<svg')
    expect(playground).toBe(github)
  })
})
