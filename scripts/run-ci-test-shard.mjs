// Runs one shard of the workspace test suite as defined in scripts/ci-test-shards.json.
//
//   node scripts/run-ci-test-shard.mjs --check          verify every workspace test script is in exactly one shard
//   node scripts/run-ci-test-shard.mjs --list           print the shard names (one per line)
//   node scripts/run-ci-test-shard.mjs <shard>          build the shard's prerequisites and run its tests
//   node scripts/run-ci-test-shard.mjs <shard> --no-build
//
// The union of all shards plus `npm run test` (which runs everything sequentially) is
// the same set of tests; the shards only exist so PR CI can run them in parallel.
import { spawnSync } from 'node:child_process'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const definition = JSON.parse(readFileSync(resolve(root, 'scripts/ci-test-shards.json'), 'utf8'))
const shards = definition.shards
const rootScripts = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8')).scripts ?? {}

const workspacesWithTests = ['apps', 'packages']
  .flatMap((base) => readdirSync(resolve(root, base), { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && existsSync(resolve(root, base, entry.name, 'package.json')))
    .map((entry) => `${base}/${entry.name}`))
  .filter((workspace) => typeof JSON.parse(readFileSync(resolve(root, workspace, 'package.json'), 'utf8')).scripts?.test === 'string')
  .sort()

function check() {
  const failures = []
  const seen = new Map()
  for (const [name, shard] of Object.entries(shards)) {
    if (!/^[a-z0-9-]+$/.test(name)) failures.push(`shard name ${JSON.stringify(name)} must be lowercase letters, digits and dashes`)
    if (!Array.isArray(shard.workspaces) || shard.workspaces.length === 0) failures.push(`shard ${name}: workspaces must be a non-empty array`)
    for (const workspace of shard.workspaces ?? []) {
      if (!workspacesWithTests.includes(workspace)) failures.push(`shard ${name}: ${workspace} is not a workspace with a test script`)
      if (seen.has(workspace)) failures.push(`${workspace} is listed in both ${seen.get(workspace)} and ${name}`)
      seen.set(workspace, name)
    }
    for (const workspace of shard.build ?? []) {
      if (!existsSync(resolve(root, workspace, 'package.json'))) failures.push(`shard ${name}: build target ${workspace} is not a workspace`)
    }
    for (const script of shard.scripts ?? []) {
      if (typeof rootScripts[script] !== 'string') failures.push(`shard ${name}: root package.json has no script ${JSON.stringify(script)}`)
    }
  }
  for (const workspace of workspacesWithTests) {
    if (!seen.has(workspace)) failures.push(`${workspace} has a test script but is not assigned to any shard in scripts/ci-test-shards.json`)
  }
  const matrix = readFileSync(resolve(root, '.github/workflows/test.yml'), 'utf8')
  for (const name of Object.keys(shards)) {
    if (!new RegExp(`^\\s+shard: \\[(?:[^\\]]*,\\s*)?${name}(?:\\s*,[^\\]]*)?\\]\\s*$`, 'm').test(matrix)) {
      failures.push(`.github/workflows/test.yml ts-tests matrix does not list shard ${name}`)
    }
  }
  if (failures.length > 0) {
    for (const failure of failures) console.error(`ci-test-shards: ${failure}`)
    process.exit(1)
  }
  console.log(`ci-test-shards: ${workspacesWithTests.length} workspace test scripts assigned across ${Object.keys(shards).length} shards`)
}

function run(command, args) {
  console.log(`\n$ ${[command, ...args].join(' ')}`)
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit', env: process.env })
  if (result.status !== 0) process.exit(result.status ?? 1)
}

const args = process.argv.slice(2)
if (args.includes('--check')) {
  check()
} else if (args.includes('--list')) {
  console.log(Object.keys(shards).join('\n'))
} else {
  const name = args.find((argument) => !argument.startsWith('--'))
  const shard = shards[name]
  if (!shard) {
    console.error(`unknown shard ${JSON.stringify(name)}; known shards: ${Object.keys(shards).join(', ')}`)
    process.exit(1)
  }
  if (!args.includes('--no-build')) {
    if (shard.build) run('npm', ['run', 'build', ...shard.build.flatMap((workspace) => ['-w', workspace])])
    else run('npm', ['run', 'build'])
  }
  run('npm', ['run', 'test', '--if-present', ...shard.workspaces.flatMap((workspace) => ['-w', workspace])])
  for (const script of shard.scripts ?? []) run('npm', ['run', script])
}
