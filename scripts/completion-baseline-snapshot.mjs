import { createHash } from 'node:crypto'
import { existsSync, readFileSync, realpathSync } from 'node:fs'
import { resolve } from 'node:path'
import { gunzipSync } from 'node:zlib'

const commit = '420424b57f658a034f7d8cdd76aab20f048a32f4'
const digest = '9fd082f8da616cc0eebc63230f8e107ad3c1a603bfe81e7fa614295145912f3b'
const snapshots = new Map()

// This immutable file snapshot preserves historical evidence in shallow clones
// and fresh-history repositories without importing any historical Git commits.
export function readCompletionBaselineSnapshot(root, baseline) {
  if (baseline !== commit) return undefined
  const key = realpathSync(root)
  if (snapshots.has(key)) return snapshots.get(key)
  const path = resolve(key, 'testdata/native-office-completion/baseline-420424b.json.gz')
  if (!existsSync(path)) return undefined
  const compressed = readFileSync(path)
  if (createHash('sha256').update(compressed).digest('hex') !== digest) {
    throw new Error('completion baseline snapshot digest does not match the pinned historical evidence')
  }
  const snapshot = JSON.parse(gunzipSync(compressed, { maxOutputLength: 32 * 1024 * 1024 }).toString('utf8'))
  if (snapshot.commit !== commit) throw new Error('completion baseline snapshot commit is incorrect')
  snapshots.set(key, snapshot)
  return snapshot
}
