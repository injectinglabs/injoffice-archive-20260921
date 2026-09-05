import { Buffer } from 'node:buffer'

export const DEPENDENCY_KINDS = [
  'dependencies',
  'peerDependencies',
  'optionalDependencies',
  'devDependencies',
]

export function packageNameFromLockPath(lockPath, entry = {}) {
  if (entry.name) return entry.name
  const marker = 'node_modules/'
  const offset = lockPath.lastIndexOf(marker)
  const suffix = offset === -1 ? lockPath : lockPath.slice(offset + marker.length)
  const parts = suffix.split('/')
  return parts[0]?.startsWith('@') ? `${parts[0]}/${parts[1]}` : parts[0]
}

export function installedNameFromLockPath(lockPath) {
  const marker = 'node_modules/'
  const offset = lockPath.lastIndexOf(marker)
  const suffix = offset === -1 ? lockPath : lockPath.slice(offset + marker.length)
  const parts = suffix.split('/')
  return parts[0]?.startsWith('@') ? `${parts[0]}/${parts[1]}` : parts[0]
}

export function integrityHash(integrity) {
  if (typeof integrity !== 'string') return null
  const match = /^(sha(?:256|384|512))-([A-Za-z0-9+/=]+)$/.exec(integrity)
  if (!match) return null
  return {
    alg: match[1].toUpperCase().replace('SHA', 'SHA-'),
    content: Buffer.from(match[2], 'base64').toString('hex'),
  }
}

export function parseGoMod(source) {
  const requires = []
  const replacements = new Map()
  let inRequire = false
  for (const rawLine of source.split(/\r?\n/)) {
    const line = rawLine.replace(/\/\/.*$/, '').trim()
    if (!line) continue
    if (line === 'require (') {
      inRequire = true
      continue
    }
    if (inRequire && line === ')') {
      inRequire = false
      continue
    }
    const requireMatch = inRequire
      ? /^(\S+)\s+(\S+)$/.exec(line)
      : /^require\s+(\S+)\s+(\S+)$/.exec(line)
    if (requireMatch) requires.push({ module: requireMatch[1], version: requireMatch[2] })
    const replaceMatch = /^replace\s+(\S+)(?:\s+\S+)?\s+=>\s+(\S+)(?:\s+(\S+))?$/.exec(line)
    if (replaceMatch) replacements.set(replaceMatch[1], { target: replaceMatch[2], version: replaceMatch[3] ?? null })
  }
  return { requires, replacements }
}

export function shortestTrace(graph, roots, target) {
  const queue = roots.map((root) => [root])
  const visited = new Set()
  while (queue.length > 0) {
    const path = queue.shift()
    const current = path.at(-1)
    if (current === target) return path
    if (visited.has(current)) continue
    visited.add(current)
    for (const next of [...(graph.get(current) ?? [])].sort()) queue.push([...path, next])
  }
  return null
}

export function purlName(name) {
  if (!name.startsWith('@')) return encodeURIComponent(name)
  const [scope, packageName] = name.split('/')
  return `${encodeURIComponent(scope)}/${encodeURIComponent(packageName)}`
}
