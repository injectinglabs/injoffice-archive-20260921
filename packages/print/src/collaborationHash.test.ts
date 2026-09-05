import { describe, expect, it } from 'vitest'
import { sha256Hex } from './collaborationHash'

describe('print collaboration SHA-256', () => {
  it.each([
    ['', 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'],
    ['abc', 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'],
    ['こんにちは', '125aeadf27b0459b8760c13a3d80912dfa8a81a68261906f60d87f4a0268646c'],
  ])('matches the published vector for %j', (input, expected) => {
    expect(sha256Hex(input)).toBe(expected)
  })
})
