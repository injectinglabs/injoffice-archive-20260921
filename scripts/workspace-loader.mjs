import { pathToFileURL } from 'node:url'
import { resolve as resolvePath } from 'node:path'

let workspaceRoot

export function initialize(data) {
  workspaceRoot = data.root
}

export async function resolve(specifier, context, nextResolve) {
  const match = /^@injoffice\/([a-z0-9-]+)$/.exec(specifier)
  if (match && workspaceRoot) {
    return {
      shortCircuit: true,
      url: pathToFileURL(resolvePath(workspaceRoot, 'packages', match[1], 'dist', 'index.js')).href,
    }
  }
  return nextResolve(specifier, context)
}
