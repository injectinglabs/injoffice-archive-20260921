import { closeSync, openSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { describe, expect, it } from 'vitest'
import { faceCoversCodepoints, fontFacesFromBuffer, isTruetype, norm, styleScore, styleTokens } from './sfnt'

function utf16be(value: string): Buffer {
  const output = Buffer.alloc(value.length * 2)
  for (let i = 0; i < value.length; i++) output.writeUInt16BE(value.charCodeAt(i), i * 2)
  return output
}

function nameTable(): Buffer {
  const names = [
    { id: 1, bytes: utf16be('Northstar Sans') },
    { id: 2, bytes: utf16be('Bold Italic') },
    { id: 6, bytes: utf16be('NorthstarSans-BoldItalic') },
  ]
  const storageStart = 6 + names.length * 12
  const size = storageStart + names.reduce((total, item) => total + item.bytes.length, 0)
  const output = Buffer.alloc(size)
  output.writeUInt16BE(0, 0)
  output.writeUInt16BE(names.length, 2)
  output.writeUInt16BE(storageStart, 4)
  let stringOffset = 0
  names.forEach((item, index) => {
    const cursor = 6 + index * 12
    output.writeUInt16BE(3, cursor)
    output.writeUInt16BE(1, cursor + 2)
    output.writeUInt16BE(0x0409, cursor + 4)
    output.writeUInt16BE(item.id, cursor + 6)
    output.writeUInt16BE(item.bytes.length, cursor + 8)
    output.writeUInt16BE(stringOffset, cursor + 10)
    item.bytes.copy(output, storageStart + stringOffset)
    stringOffset += item.bytes.length
  })
  return output
}

function cmapTable(): Buffer {
  const output = Buffer.alloc(12 + 32)
  output.writeUInt16BE(0, 0)
  output.writeUInt16BE(1, 2)
  output.writeUInt16BE(3, 4)
  output.writeUInt16BE(1, 6)
  output.writeUInt32BE(12, 8)
  const base = 12
  output.writeUInt16BE(4, base)
  output.writeUInt16BE(32, base + 2)
  output.writeUInt16BE(0, base + 4)
  output.writeUInt16BE(4, base + 6)
  output.writeUInt16BE(4, base + 8)
  output.writeUInt16BE(1, base + 10)
  output.writeUInt16BE(0, base + 12)
  output.writeUInt16BE(65, base + 14)
  output.writeUInt16BE(0xffff, base + 16)
  output.writeUInt16BE(0, base + 18)
  output.writeUInt16BE(65, base + 20)
  output.writeUInt16BE(0xffff, base + 22)
  output.writeUInt16BE(0xffc0, base + 24)
  output.writeUInt16BE(1, base + 26)
  output.writeUInt16BE(0, base + 28)
  output.writeUInt16BE(0, base + 30)
  return output
}

function syntheticTrueType(): Buffer {
  const tables = [
    { tag: 'cmap', bytes: cmapTable() },
    { tag: 'glyf', bytes: Buffer.alloc(0) },
    { tag: 'head', bytes: Buffer.alloc(12) },
    { tag: 'loca', bytes: Buffer.alloc(0) },
    { tag: 'name', bytes: nameTable() },
  ]
  const directorySize = 12 + tables.length * 16
  const total = directorySize + tables.reduce((size, table) => size + ((table.bytes.length + 3) & ~3), 0)
  const output = Buffer.alloc(total)
  output.writeUInt32BE(0x00010000, 0)
  output.writeUInt16BE(tables.length, 4)
  let dataOffset = directorySize
  tables.forEach((table, index) => {
    const cursor = 12 + index * 16
    output.write(table.tag, cursor, 4, 'latin1')
    output.writeUInt32BE(dataOffset, cursor + 8)
    output.writeUInt32BE(table.bytes.length, cursor + 12)
    table.bytes.copy(output, dataOffset)
    dataOffset += (table.bytes.length + 3) & ~3
  })
  return output
}

describe('independently authored OpenType fixtures', () => {
  it('normalizes names and compares semantic style coordinates', () => {
    expect(norm(' Nórthstar-Sans ')).toBe('northstarsans')
    expect(styleTokens('NorthstarSans-BoldItalic')).toEqual(['bold', 'italic'])
    const face = { path: '/fixture.ttf', offset: 0, style: 'Bold Italic' }
    expect(styleScore(face, ['bold', 'italic'])).toBeGreaterThan(styleScore(face, []))
  })

  it('reads Windows Unicode names and rejects a CFF signature as TrueType outlines', () => {
    const bytes = syntheticTrueType()
    expect(fontFacesFromBuffer(bytes, '/fixture.ttf')).toEqual([{
      ref: { path: '/fixture.ttf', offset: 0, style: 'Bold Italic' },
      names: {
        ps: ['NorthstarSans-BoldItalic'],
        families: ['Northstar Sans'],
        subfamilies: ['Bold Italic'],
      },
    }])
    expect(isTruetype(bytes)).toBe(true)
    const cff = Buffer.from(bytes)
    cff.write('OTTO', 0, 4, 'latin1')
    expect(isTruetype(cff)).toBe(false)
  })

  it('uses nonzero glyph mappings for cmap coverage', () => {
    const path = join(tmpdir(), `injoffice-northstar-${process.pid}-${Date.now()}.ttf`)
    writeFileSync(path, syntheticTrueType())
    const fd = openSync(path, 'r')
    try {
      expect(faceCoversCodepoints(fd, 0, [65])).toBe(true)
      expect(faceCoversCodepoints(fd, 0, [66])).toBe(false)
    } finally {
      closeSync(fd)
    }
  })

  it('fails closed on malformed table offsets', () => {
    const bytes = syntheticTrueType()
    bytes.writeUInt32BE(0xfffffff0, 12 + 8)
    expect(fontFacesFromBuffer(bytes, '/broken.ttf')).toEqual([])
  })
})
