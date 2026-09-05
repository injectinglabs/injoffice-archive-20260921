import { copyFile, mkdir, readFile, rm } from 'node:fs/promises'
import { basename, resolve } from 'node:path'
import { builtinModules } from 'node:module'
import { build } from 'vite'

const directory = resolve(process.cwd())
const manifest = JSON.parse(await readFile(resolve(directory, 'package.json'), 'utf8'))
const entries = manifest.injoffice?.entries ?? { index: 'src/index.ts' }
const externalPackages = new Set([
  ...Object.keys(manifest.dependencies ?? {}),
  ...Object.keys(manifest.peerDependencies ?? {}),
  ...builtinModules,
  ...builtinModules.map((name) => `node:${name}`),
])

const isExternal = (id) => {
  for (const name of externalPackages) {
    if (id === name || id.startsWith(`${name}/`)) return true
  }
  return false
}

if (manifest.injoffice?.tscOnly) {
  await rm(resolve(directory, 'dist'), { recursive: true, force: true })
  await mkdir(resolve(directory, 'dist'), { recursive: true })
} else {
  await build({
    configFile: false,
    root: directory,
    logLevel: 'warn',
    build: {
      emptyOutDir: true,
      lib: {
        entry: Object.fromEntries(Object.entries(entries).map(([name, path]) => [name, resolve(directory, path)])),
        formats: ['es'],
        fileName: (_format, name) => `${name}.js`,
      },
      minify: false,
      sourcemap: true,
      rollupOptions: { external: isExternal },
    },
  })
}

const workspaceRoot = resolve(directory, '../..')
await copyFile(resolve(workspaceRoot, 'LICENSE'), resolve(directory, 'dist', 'LICENSE'))
for (const legalFile of manifest.injoffice?.legalFiles ?? []) {
  await copyFile(resolve(workspaceRoot, legalFile), resolve(directory, 'dist', basename(legalFile)))
}
