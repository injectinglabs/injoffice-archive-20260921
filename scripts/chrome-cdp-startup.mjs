import { spawn } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const DEFAULT_ATTEMPTS = 3
const DEFAULT_TIMEOUT_MS = 15_000

export async function launchChromeForCDP({
  executable,
  createProfile,
  attempts = DEFAULT_ATTEMPTS,
  timeoutMs = DEFAULT_TIMEOUT_MS,
  spawnProcess = spawn,
  fetchImpl = fetch,
  readActivePort = readDevToolsActivePort,
  sleep = delay,
  now = Date.now,
  stopProcess = terminateProcess,
  onAttemptFailure = () => {},
}) {
  if (!Number.isInteger(attempts) || attempts < 1) throw new Error('Chrome startup attempts must be a positive integer')
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new Error('Chrome startup timeout must be positive')

  const failures = []
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    const startedAt = now()
    const attemptTimeoutMs = timeoutMs * attempt
    const profile = createProfile(attempt)
    let stderr = ''
    let spawnError
    let child
    try {
      child = spawnProcess(executable, chromeArguments(profile), { stdio: ['ignore', 'ignore', 'pipe'] })
      child.once('error', (error) => { spawnError = error })
      child.stderr?.setEncoding('utf8')
      child.stderr?.on('data', (chunk) => { stderr = `${stderr}${chunk}`.slice(-8_000) })
      const target = await waitForTarget({
        profile,
        child,
        timeoutMs: attemptTimeoutMs,
        fetchImpl,
        readActivePort,
        sleep,
        now,
        getSpawnError: () => spawnError,
      })
      return {
        child,
        target,
        profile,
        startupFailures: failures,
        diagnostics: () => formatDiagnostics(failures, {
          attempt,
          message: 'Chrome exposed a debugging target',
          elapsedMs: Math.max(0, now() - startedAt),
          process: processStatus(child),
          stderr,
        }),
      }
    } catch (error) {
      const failure = {
        attempt,
        message: error instanceof Error ? error.message : String(error),
        elapsedMs: Math.max(0, now() - startedAt),
        process: processStatus(child),
        stderr,
      }
      failures.push(failure)
      if (child && isProcessRunning(child)) {
        try {
          await stopProcess(child)
        } catch (stopError) {
          failure.stopError = stopError instanceof Error ? stopError.message : String(stopError)
          throw new Error(`Chrome startup cleanup failed; refusing to launch another process:\n${formatDiagnostics(failures)}`, {
            cause: stopError,
          })
        }
      }
      onAttemptFailure(failure, { attempt, attempts })
      if (attempt === attempts) {
        throw new Error(`Chrome did not expose a debugging target after ${attempts} attempts:\n${formatDiagnostics(failures)}`, {
          cause: error,
        })
      }
      await sleep(Math.min(250 * attempt, 1_000))
    }
  }
  throw new Error('unreachable Chrome startup state')
}

export function chromeArguments(profile) {
  return [
    '--headless=new',
    '--no-sandbox',
    '--disable-dev-shm-usage',
    '--disable-gpu',
    '--disable-background-networking',
    '--disable-component-update',
    '--disable-extensions',
    '--no-first-run',
    '--no-default-browser-check',
    '--remote-allow-origins=*',
    '--remote-debugging-address=127.0.0.1',
    '--remote-debugging-port=0',
    `--user-data-dir=${profile}`,
    'about:blank',
  ]
}

export function readDevToolsActivePort(profile) {
  const contents = readFileSync(resolve(profile, 'DevToolsActivePort'), 'utf8')
  const [line] = contents.split(/\r?\n/)
  const port = Number(line)
  if (!Number.isInteger(port) || port < 1 || port > 65_535) {
    throw new Error(`DevToolsActivePort contained an invalid port: ${JSON.stringify(line)}`)
  }
  return port
}

export async function waitForTarget({
  profile,
  child,
  timeoutMs,
  fetchImpl,
  readActivePort,
  sleep,
  now,
  getSpawnError = () => undefined,
}) {
  const deadline = now() + timeoutMs
  let activePortState = 'not created'
  let lastProbeError = 'not attempted'
  while (now() < deadline) {
    const spawnError = getSpawnError()
    if (spawnError) {
      throw new Error(`Chrome process failed to start: ${spawnError instanceof Error ? spawnError.message : String(spawnError)}`)
    }
    if (!isProcessRunning(child)) {
      throw new Error(`Chrome exited before exposing its debugging target (${processStatus(child)})`)
    }
    let port
    try {
      port = readActivePort(profile)
      activePortState = String(port)
    } catch (error) {
      activePortState = error?.code === 'ENOENT'
        ? 'not created'
        : `unreadable (${error instanceof Error ? error.message : String(error)})`
    }
    if (port !== undefined) {
      try {
        const remaining = Math.max(1, deadline - now())
        const response = await fetchImpl(`http://127.0.0.1:${port}/json/list`, {
          signal: AbortSignal.timeout(Math.min(1_000, remaining)),
        })
        if (!response.ok) throw new Error(`HTTP ${response.status}`)
        const targets = await response.json()
        if (!Array.isArray(targets)) throw new Error('target list was not an array')
        const target = targets.find((item) => item.type === 'page')
        if (target?.webSocketDebuggerUrl) return target
        lastProbeError = 'no page target was present'
      } catch (error) {
        lastProbeError = error instanceof Error ? error.message : String(error)
      }
    }
    await sleep(50)
  }
  throw new Error(
    `timed out after ${timeoutMs}ms waiting for the Chrome debugging target `
    + `(DevToolsActivePort: ${activePortState}; last probe: ${lastProbeError})`,
  )
}

export async function terminateProcess(child, timeoutMs = 3_000) {
  if (!isProcessRunning(child)) return
  let exited = waitForExit(child)
  child.kill('SIGTERM')
  await Promise.race([exited, delay(timeoutMs)])
  if (!isProcessRunning(child)) return
  exited = waitForExit(child)
  child.kill('SIGKILL')
  await Promise.race([exited, delay(timeoutMs)])
  if (isProcessRunning(child)) throw new Error(`Chrome remained running after SIGKILL (${processStatus(child)})`)
}

function isProcessRunning(child) {
  return child != null && child.exitCode === null && child.signalCode === null
}

function processStatus(child) {
  if (!child) return 'not spawned'
  if (child.exitCode !== null) return `exit code ${child.exitCode}`
  if (child.signalCode !== null) return `signal ${child.signalCode}`
  return `running (pid ${child.pid ?? 'unknown'})`
}

function waitForExit(child) {
  return new Promise((resolvePromise) => child.once('exit', resolvePromise))
}

function formatDiagnostics(failures, current) {
  return [...failures, ...(current ? [current] : [])].map((failure) => {
    const fields = [
      `attempt ${failure.attempt}: ${failure.message}`,
      `elapsed: ${failure.elapsedMs}ms`,
      `process: ${failure.process}`,
      ...(failure.stopError ? [`cleanup: ${failure.stopError}`] : []),
      `stderr: ${failure.stderr.trim() || '(empty)'}`,
    ]
    return fields.join('\n  ')
  }).join('\n')
}

function delay(milliseconds) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, milliseconds))
}
