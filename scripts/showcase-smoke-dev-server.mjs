import { basename, dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

/** Only the three installed WASM distributions, never the linked checkout root. */
export function showcaseWasmDistDirectories(resolveModule = (specifier) => import.meta.resolve(specifier)) {
  return ['xlsx', 'docx', 'pptx'].map((format) => {
    const distribution = dirname(fileURLToPath(resolveModule(`@injoffice/${format}-wasm`)))
    if (basename(distribution) !== 'dist' || basename(dirname(distribution)) !== `${format}-wasm`) {
      throw new Error(`Unexpected @injoffice/${format}-wasm distribution path; refusing to widen the smoke filesystem allowance.`)
    }
    return distribution
  })
}

/** Keep development React diagnostics without binding either user's port 3100. */
export function isolatedShowcaseDevConfig(config, root, { workspaceRoot = root, wasmDirectories = [] } = {}) {
  return {
    ...config,
    configFile: false,
    root,
    plugins: (config.plugins ?? []).filter((plugin) => plugin?.name !== 'injoffice-ipv6-loopback'),
    server: {
      ...config.server, host: '127.0.0.1', port: 0, strictPort: false, open: false,
      fs: {
        ...config.server?.fs,
        allow: [...new Set([...(config.server?.fs?.allow ?? [workspaceRoot]), ...wasmDirectories])],
      },
    },
  }
}

export async function startShowcaseDevServer(root) {
  const { createServer, loadConfigFromFile, searchForWorkspaceRoot } = await import('vite')
  const loaded = await loadConfigFromFile({ command: 'serve', mode: 'development' }, resolve(root, 'vite.config.ts'))
  if (!loaded) throw new Error('The playground Vite configuration could not be loaded.')
  const server = await createServer(isolatedShowcaseDevConfig(loaded.config, root, {
    workspaceRoot: searchForWorkspaceRoot(root), wasmDirectories: showcaseWasmDistDirectories(),
  }))
  try {
    await server.listen()
    return {
      url: `http://127.0.0.1:${server.httpServer.address().port}/`,
      close: () => server.close(),
    }
  } catch (error) {
    await server.close()
    throw error
  }
}
