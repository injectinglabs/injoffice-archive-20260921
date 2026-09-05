import assert from 'node:assert/strict'
import { EventEmitter } from 'node:events'
import test from 'node:test'

import {
  chromeArguments,
  launchChromeForCDP,
  waitForTarget,
} from './chrome-cdp-startup.mjs'

test('Chrome chooses its debugging port and uses a dedicated profile', () => {
  const args = chromeArguments('/tmp/injoffice-profile')
  assert.ok(args.includes('--remote-debugging-port=0'))
  assert.ok(args.includes('--remote-debugging-address=127.0.0.1'))
  assert.ok(args.includes('--user-data-dir=/tmp/injoffice-profile'))
  assert.ok(args.includes('--no-first-run'))
})

test('startup retries with a fresh profile and stops the failed process', async () => {
  const children = [fakeChild(101), fakeChild(102)]
  const profiles = []
  const stopped = []
  const warnings = []
  let launches = 0
  const result = await launchChromeForCDP({
    executable: 'chrome',
    createProfile(attempt) {
      const profile = `/tmp/profile-${attempt}`
      profiles.push(profile)
      return profile
    },
    spawnProcess() {
      const child = children[launches]
      launches += 1
      return child
    },
    readActivePort(profile) {
      if (profile.endsWith('-1')) throw Object.assign(new Error('missing'), { code: 'ENOENT' })
      return 43210
    },
    fetchImpl: async () => ({
      ok: true,
      json: async () => [{ type: 'page', webSocketDebuggerUrl: 'ws://target' }],
    }),
    timeoutMs: 100,
    now: incrementingClock(60),
    sleep: async () => {},
    stopProcess: async (child) => {
      stopped.push(child.pid)
      child.signalCode = 'SIGTERM'
    },
    onAttemptFailure: (failure) => warnings.push(failure),
  })

  assert.deepEqual(profiles, ['/tmp/profile-1', '/tmp/profile-2'])
  assert.deepEqual(stopped, [101])
  assert.equal(warnings.length, 1)
  assert.equal(result.child.pid, 102)
  assert.equal(result.target.webSocketDebuggerUrl, 'ws://target')
  assert.match(result.diagnostics(), /attempt 1:/)
  assert.match(result.diagnostics(), /attempt 2: Chrome exposed a debugging target/)
})

test('startup reports bounded per-attempt diagnostics after exhausting retries', async () => {
  const children = [fakeChild(201), fakeChild(202)]
  const emittedStderr = []
  let launches = 0
  await assert.rejects(
    launchChromeForCDP({
      executable: 'chrome',
      createProfile: (attempt) => `/tmp/profile-${attempt}`,
      spawnProcess() {
        const child = children[launches]
        launches += 1
        const stderr = `old-prefix-${launches}-${'x'.repeat(9_000)}-new-suffix-${launches}`
        emittedStderr.push(stderr)
        queueMicrotask(() => child.stderr.emit('data', stderr))
        return child
      },
      readActivePort: () => { throw Object.assign(new Error('missing'), { code: 'ENOENT' }) },
      fetchImpl: async () => { throw new Error('must not fetch without a port') },
      attempts: 2,
      timeoutMs: 100,
      now: incrementingClock(60),
      sleep: async () => {},
      stopProcess: async (child) => { child.signalCode = 'SIGTERM' },
    }),
    (error) => {
      assert.match(error.message, /after 2 attempts/)
      assert.match(error.message, /attempt 1:/)
      assert.match(error.message, /attempt 2:/)
      assert.match(error.message, /timed out after 100ms/)
      assert.match(error.message, /timed out after 200ms/)
      assert.match(error.message, /elapsed: \d+ms/)
      assert.doesNotMatch(error.message, /old-prefix-/)
      assert.match(error.message, /new-suffix-1/)
      assert.match(error.message, /new-suffix-2/)
      return true
    },
  )
  assert.equal(launches, 2)
  assert.ok(emittedStderr.every((stderr) => stderr.length > 8_000))
})

test('target polling fails immediately when Chrome exits', async () => {
  const child = fakeChild(301)
  child.exitCode = 17
  await assert.rejects(
    waitForTarget({
      profile: '/tmp/profile',
      child,
      timeoutMs: 100,
      fetchImpl: async () => { throw new Error('unreachable') },
      readActivePort: () => 1234,
      sleep: async () => {},
      now: Date.now,
    }),
    /exit code 17/,
  )
})

test('an asynchronous spawn error is captured instead of becoming an uncaught event', async () => {
  const child = fakeChild(401)
  let spawnError
  await assert.rejects(
    waitForTarget({
      profile: '/tmp/profile',
      child,
      timeoutMs: 100,
      fetchImpl: async () => { throw new Error('unreachable') },
      readActivePort: () => { throw Object.assign(new Error('missing'), { code: 'ENOENT' }) },
      sleep: async () => { spawnError = new Error('permission denied') },
      now: incrementingClock(10),
      getSpawnError: () => spawnError,
    }),
    /process failed to start: permission denied/,
  )
})

function fakeChild(pid) {
  const child = new EventEmitter()
  child.pid = pid
  child.exitCode = null
  child.signalCode = null
  child.stderr = new EventEmitter()
  child.stderr.setEncoding = () => {}
  child.kill = () => true
  return child
}

function incrementingClock(step) {
  let current = 0
  return () => {
    current += step
    return current
  }
}
