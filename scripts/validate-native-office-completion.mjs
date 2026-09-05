import { createHash } from 'node:crypto'
import { spawnSync } from 'node:child_process'
import { existsSync, lstatSync, readFileSync, readdirSync, realpathSync, statSync } from 'node:fs'
import { basename, dirname, extname, isAbsolute, posix, relative, resolve, sep } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

export const PROTOCOL = 'injoffice.native-office-completion/v1'
export const PROTOCOL_V2 = 'injoffice.native-office-completion/v2'
export const PROTOCOL_V3 = 'injoffice.native-office-completion/v3'
export const BASELINE = '0cdd7e9340e9469326f57e6175ec87e7e1807443'
export const BASELINE_V2 = '0cdd7e9340e9469326f57e6175ec87e7e1807443'
export const BASELINE_V3 = '0cdd7e9340e9469326f57e6175ec87e7e1807443'
const baselineByVersion = { 1: BASELINE, 2: BASELINE_V2, 3: BASELINE_V3 }
const protocolByVersion = { 1: PROTOCOL, 2: PROTOCOL_V2, 3: PROTOCOL_V3 }
let activeBaseline = BASELINE
export const FORMATS = ['docx', 'pptx', 'xlsx']
export const GATES = ['structuralPreservation', 'renderLayout', 'mutationRoundtrip', 'packageConsumer', 'unsupportedObjectLoss', 'productionE2E']
export const DIMENSIONS = ['native', 'package-preservation', 'identity', 'security', 'resource', 'visual-production']
export const CAPABILITY_DIMENSIONS = [...DIMENSIONS, 'mutation', 'package-consumer'].sort()
export const PRODUCTION_E2E_EVIDENCE_PATHS = [
  'scripts/verify-native-office-production-e2e.mjs',
  'scripts/verify-native-office-production-e2e.test.mjs',
  'testdata/native-office-production-e2e/v1/manifest.json',
]

const scriptRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const idPattern = /^[a-z0-9]+(?:[.-][a-z0-9]+)*$/
const capabilityPattern = /^(docx|pptx|xlsx)\.[a-z0-9]+(?:[.-][a-z0-9]+)*$/
const shaPattern = /^[a-f0-9]{64}$/
const baselineFileCache = new Map()
const baselineTreeCache = new Map()
const forbiddenAuthority = [
  ['Mammoth', /(?:(?:\bfrom\s*|\bimport\s*(?:\(\s*)?|\brequire\s*\(\s*)['"]mammoth(?:\/[^'"]*)?['"])/i],
  ['Konva', /(?:(?:\bfrom\s*|\bimport\s*(?:\(\s*)?|\brequire\s*\(\s*)['"](?:react-)?konva['"])/i],
  ['legacy DeckView', /(?:\b(?:import|export)\s*\{[^}]*\bDeck(?:Canvas)?View\b|\bfrom\s*['"][^'"]*Deck(?:Canvas)?View|<Deck(?:Canvas)?View\b)/],
  ['DOM parser', /\bDOMParser\b/],
  ['DOM creation', /\bdocument\s*\.\s*createElement\b/],
  ['DOM HTML authority', /\b(?:innerHTML|outerHTML|dangerouslySetInnerHTML)\b/],
  ['HTML renderer authority', /(?:(?:\bfrom\s*|\bimport\s*(?:\(\s*)?|\brequire\s*\(\s*)['"]react-dom(?:\/[^'"]*)?['"]|\bReactDOM\b)/],
  ['browser global authority', /\bwindow\s*\.\s*(?:document|getComputedStyle)\b/],
]

const object = (value) => value !== null && typeof value === 'object' && !Array.isArray(value)
const ownKeys = (value) => object(value) ? Object.keys(value) : []
const canonicalSha = (bytes) => createHash('sha256').update(bytes).digest('hex')
const sortedUnique = (values) => Array.isArray(values) && values.every((value, index) => typeof value === 'string' && value.length > 0 && (index === 0 || values[index - 1] < value))

function schemaTypeMatches(type, value) {
  if (type === 'object') return object(value)
  if (type === 'array') return Array.isArray(value)
  if (type === 'string') return typeof value === 'string'
  if (type === 'integer') return Number.isInteger(value)
  if (type === 'number') return typeof value === 'number' && Number.isFinite(value)
  if (type === 'boolean') return typeof value === 'boolean'
  if (type === 'null') return value === null
  return true
}

function resolveSchemaReference(rootSchema, reference) {
  if (!reference.startsWith('#/')) throw new Error(`external schema reference is forbidden: ${reference}`)
  return reference.slice(2).split('/').reduce((value, key) => value?.[key.replaceAll('~1', '/').replaceAll('~0', '~')], rootSchema)
}

function validateSchemaNode(schema, value, rootSchema, path, errors) {
  if (schema.$ref) {
    const target = resolveSchemaReference(rootSchema, schema.$ref)
    if (!target) errors.push(`${path} has unresolved schema reference ${schema.$ref}`)
    else validateSchemaNode(target, value, rootSchema, path, errors)
    return
  }
  if (schema.type && !schemaTypeMatches(schema.type, value)) {
    errors.push(`${path} must have schema type ${schema.type}`)
    return
  }
  if ('const' in schema && JSON.stringify(value) !== JSON.stringify(schema.const)) errors.push(`${path} must equal ${JSON.stringify(schema.const)}`)
  if (schema.enum && !schema.enum.some((candidate) => JSON.stringify(candidate) === JSON.stringify(value))) errors.push(`${path} is not in the schema enum`)
  if (typeof value === 'string') {
    if (schema.minLength !== undefined && value.length < schema.minLength) errors.push(`${path} is shorter than schema minLength ${schema.minLength}`)
    if (schema.pattern && !new RegExp(schema.pattern).test(value)) errors.push(`${path} does not match schema pattern ${schema.pattern}`)
  }
  if (typeof value === 'number' && schema.minimum !== undefined && value < schema.minimum) errors.push(`${path} is below schema minimum ${schema.minimum}`)
  if (Array.isArray(value)) {
    if (schema.minItems !== undefined && value.length < schema.minItems) errors.push(`${path} has fewer than schema minItems ${schema.minItems}`)
    if (schema.maxItems !== undefined && value.length > schema.maxItems) errors.push(`${path} has more than schema maxItems ${schema.maxItems}`)
    if (schema.uniqueItems && new Set(value.map((entry) => JSON.stringify(entry))).size !== value.length) errors.push(`${path} violates schema uniqueItems`)
    if (schema.items) value.forEach((entry, index) => validateSchemaNode(schema.items, entry, rootSchema, `${path}[${index}]`, errors))
  }
  if (object(value)) {
    for (const key of schema.required ?? []) if (!(key in value)) errors.push(`${path} is missing schema-required field ${key}`)
    if (schema.additionalProperties === false) for (const key of Object.keys(value)) if (!(key in (schema.properties ?? {}))) errors.push(`${path} has schema-forbidden field ${key}`)
    for (const [key, childSchema] of Object.entries(schema.properties ?? {})) if (key in value) validateSchemaNode(childSchema, value[key], rootSchema, `${path}.${key}`, errors)
  }
}

function validatePublishedSchema(root, manifest) {
  const errors = []
  const version = [1, 2, 3].includes(manifest?.version) ? manifest.version : 1
  const path = resolve(root, `schemas/native-office-completion-v${version}.schema.json`)
  try {
    const schema = JSON.parse(readFileSync(path, 'utf8'))
    validateSchemaNode(schema, manifest, schema, '$', errors)
  } catch (error) {
    errors.push(`published completion schema cannot be applied: ${error instanceof Error ? error.message : String(error)}`)
  }
  return errors
}

function exactKeys(value, required, context, errors, optional = []) {
  if (!object(value)) {
    errors.push(`${context} must be an object`)
    return false
  }
  const allowed = new Set([...required, ...optional])
  for (const key of ownKeys(value)) if (!allowed.has(key)) errors.push(`${context} has unknown field ${key}`)
  for (const key of required) if (!(key in value)) errors.push(`${context} is missing ${key}`)
  return true
}

function containedPath(root, path) {
  const relativePath = relative(root, path)
  return relativePath === '' || (!isAbsolute(relativePath) && relativePath !== '..' && !relativePath.startsWith(`..${sep}`))
}

function safePath(root, value, context, errors) {
  if (typeof value !== 'string' || value.length === 0 || value.includes('\\') || value.startsWith('/') || value.split('/').some((part) => part === '' || part === '.' || part === '..')) {
    errors.push(`${context} is not a safe repository-relative path`)
    return undefined
  }
  const path = resolve(root, value)
  if (!containedPath(resolve(root), path)) {
    errors.push(`${context} escapes the repository`)
    return undefined
  }
  if (!existsSync(path) || !statSync(path).isFile()) {
    errors.push(`${context} is missing: ${value}`)
    return undefined
  }
  if (lstatSync(path).isSymbolicLink()) {
    errors.push(`${context} must not be a symbolic link: ${value}`)
    return undefined
  }
  const realRoot = realpathSync(root)
  const realPath = realpathSync(path)
  if (!containedPath(realRoot, realPath)) {
    errors.push(`${context} resolves outside the repository: ${value}`)
    return undefined
  }
  return path
}

function readBaselineFile(root, repositoryPath, context, errors) {
  const cacheKey = `${realpathSync(root)}\0${activeBaseline}\0${repositoryPath}`
  let result = baselineFileCache.get(cacheKey)
  if (!result) {
    result = spawnSync('git', ['-C', root, 'show', `${activeBaseline}:${repositoryPath}`], { encoding: null, maxBuffer: 128 * 1024 * 1024 })
    baselineFileCache.set(cacheKey, result)
  }
  if (result.status !== 0) {
    errors.push(`${context} is absent from baseline ${activeBaseline}: ${repositoryPath}`)
    return undefined
  }
  return result.stdout
}

function readBaselineTree(root) {
  const cacheKey = `${realpathSync(root)}\0${activeBaseline}`
  let tree = baselineTreeCache.get(cacheKey)
  if (tree) return tree
  const result = spawnSync('git', ['-C', root, 'ls-tree', '-r', '-z', activeBaseline], { encoding: 'utf8', maxBuffer: 128 * 1024 * 1024 })
  if (result.status !== 0) throw new Error(`cannot read baseline tree ${activeBaseline}: ${result.stderr?.trim() || `git exited ${result.status}`}`)
  tree = new Map()
  for (const record of result.stdout.split('\0')) {
    if (!record) continue
    const match = /^(\d+) (\w+) ([a-f0-9]+)\t(.+)$/.exec(record)
    if (!match) throw new Error(`invalid git ls-tree record for baseline ${activeBaseline}`)
    tree.set(match[4], { mode: match[1], type: match[2], object: match[3] })
  }
  baselineTreeCache.set(cacheKey, tree)
  return tree
}

function baselineSourcePath(tree, candidate) {
  const normalized = posix.normalize(candidate)
  if (posix.isAbsolute(normalized) || normalized === '..' || normalized.startsWith('../')) return undefined
  const entry = tree.get(normalized)
  return entry?.type === 'blob' && (entry.mode === '100644' || entry.mode === '100755') ? normalized : undefined
}

function requireStrings(value, context, errors, { nonempty = true, sorted = false } = {}) {
  if (!Array.isArray(value) || (nonempty && value.length === 0) || value.some((entry) => typeof entry !== 'string' || entry.length === 0)) {
    errors.push(`${context} must be ${nonempty ? 'a non-empty' : 'an'} string array`)
    return false
  }
  if ((sorted && !sortedUnique(value)) || new Set(value).size !== value.length) errors.push(`${context} must be sorted and unique`)
  return true
}

export function findForbiddenNativeAuthority(source) {
  return forbiddenAuthority.filter(([, pattern]) => pattern.test(source)).map(([label]) => label)
}

// Replace actual comments with whitespace while preserving strings and template
// literals. Deleting comments joins tokens (`from/* gap */'pkg'`), while a
// regex-only stripper can be poisoned by comment markers inside strings.
export function normalizeSourceComments(source) {
  let output = ''
  let index = 0
  let mode = 'code'
  const templateExpressions = []
  const whitespace = (character) => character === '\n' || character === '\r' ? character : ' '
  let regexCharacterClass = false
  const regexCanStart = () => {
    const significant = output.trimEnd()
    if (significant.length === 0) return true
    if ('([{=,:;!&|?+-*%^~<>'.includes(significant.at(-1))) return true
    return /\b(?:await|case|delete|do|else|in|instanceof|of|return|throw|typeof|void|yield)$/.test(significant)
  }

  while (index < source.length) {
    const character = source[index]
    const next = source[index + 1]

    if (mode === 'single' || mode === 'double') {
      output += character
      index += 1
      if (character === '\\' && index < source.length) {
        output += source[index]
        index += 1
      } else if ((mode === 'single' && character === "'") || (mode === 'double' && character === '"')) {
        mode = 'code'
      }
      continue
    }

    if (mode === 'template') {
      output += character
      index += 1
      if (character === '\\' && index < source.length) {
        output += source[index]
        index += 1
      } else if (character === '`') {
        mode = 'code'
      } else if (character === '$' && next === '{') {
        output += next
        index += 1
        templateExpressions.push(0)
        mode = 'code'
      }
      continue
    }

    if (mode === 'regex') {
      output += character
      index += 1
      if (character === '\\' && index < source.length) {
        output += source[index]
        index += 1
      } else if (character === '[') {
        regexCharacterClass = true
      } else if (character === ']') {
        regexCharacterClass = false
      } else if (character === '/' && !regexCharacterClass) {
        mode = 'code'
      }
      continue
    }

    if (character === '/' && next === '/') {
      output += '  '
      index += 2
      while (index < source.length && source[index] !== '\n' && source[index] !== '\r') {
        output += ' '
        index += 1
      }
      continue
    }
    if (character === '/' && next === '*') {
      output += '  '
      index += 2
      while (index < source.length) {
        if (source[index] === '*' && source[index + 1] === '/') {
          output += '  '
          index += 2
          break
        }
        output += whitespace(source[index])
        index += 1
      }
      continue
    }
    if (character === '/' && regexCanStart()) {
      mode = 'regex'
      regexCharacterClass = false
      output += character
      index += 1
      continue
    }
    if (character === "'") mode = 'single'
    else if (character === '"') mode = 'double'
    else if (character === '`') mode = 'template'
    else if (templateExpressions.length > 0 && character === '{') templateExpressions[templateExpressions.length - 1] += 1
    else if (templateExpressions.length > 0 && character === '}') {
      if (templateExpressions[templateExpressions.length - 1] === 0) {
        templateExpressions.pop()
        mode = 'template'
      } else templateExpressions[templateExpressions.length - 1] -= 1
    }
    output += character
    index += 1
  }
  return output
}

function staticSourceDependencies(source) {
  const specifiers = new Set()
  const patterns = [
    /\b(?:import|export)\s+(?:type\s+)?[^'"`;]*?\bfrom\s*(['"])([^'"\r\n]+)\1/g,
    /\bimport\s*(['"])([^'"\r\n]+)\1/g,
    /\bimport\s*\(\s*(['"])([^'"\r\n]+)\1\s*\)/g,
    /\brequire\s*\(\s*(['"])([^'"\r\n]+)\1\s*\)/g,
  ]
  for (const pattern of patterns) for (const match of source.matchAll(pattern)) specifiers.add(match[2])
  return [...specifiers].sort()
}

function resolveSourceImport(root, importer, specifier) {
  const candidates = []
  if (specifier.startsWith('.')) {
    const unresolved = resolve(dirname(importer), specifier)
    candidates.push(unresolved)
    if (['.js', '.jsx', '.mjs', '.cjs'].includes(extname(unresolved))) candidates.push(unresolved.replace(/\.[^.]+$/, '.ts'), unresolved.replace(/\.[^.]+$/, '.tsx'), unresolved.replace(/\.[^.]+$/, '.mts'), unresolved.replace(/\.[^.]+$/, '.cts'))
    else candidates.push(`${unresolved}.js`, `${unresolved}.jsx`, `${unresolved}.mjs`, `${unresolved}.cjs`, `${unresolved}.ts`, `${unresolved}.tsx`, `${unresolved}.mts`, `${unresolved}.cts`, resolve(unresolved, 'index.js'), resolve(unresolved, 'index.ts'))
  } else if (specifier.startsWith('@injoffice/')) {
    const segments = specifier.split('/')
    const packageName = segments.slice(0, 2).join('/')
    const subpath = segments.slice(2).join('/') || 'index'
    const packagesRoot = resolve(root, 'packages')
    if (existsSync(packagesRoot)) {
      for (const directory of statDirectoryNames(packagesRoot)) {
        const manifestPath = resolve(packagesRoot, directory, 'package.json')
        if (!existsSync(manifestPath)) continue
        const packageManifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
        if (packageManifest.name !== packageName) continue
        const entry = packageManifest.injoffice?.entries?.[subpath] ?? (subpath === 'index' ? 'src/index.ts' : `src/${subpath}.ts`)
        candidates.push(resolve(packagesRoot, directory, entry))
      }
    }
  }
  return candidates.find((candidate) => existsSync(candidate) && statSync(candidate).isFile())
}

function resolveBaselineSourceImport(root, tree, importer, specifier) {
  const candidates = []
  if (specifier.startsWith('.')) {
    const unresolved = posix.normalize(posix.join(posix.dirname(importer), specifier))
    candidates.push(unresolved)
    if (['.js', '.jsx', '.mjs', '.cjs'].includes(posix.extname(unresolved))) candidates.push(unresolved.replace(/\.[^.]+$/, '.ts'), unresolved.replace(/\.[^.]+$/, '.tsx'), unresolved.replace(/\.[^.]+$/, '.mts'), unresolved.replace(/\.[^.]+$/, '.cts'))
    else candidates.push(`${unresolved}.js`, `${unresolved}.jsx`, `${unresolved}.mjs`, `${unresolved}.cjs`, `${unresolved}.ts`, `${unresolved}.tsx`, `${unresolved}.mts`, `${unresolved}.cts`, posix.join(unresolved, 'index.js'), posix.join(unresolved, 'index.ts'))
  } else if (specifier.startsWith('@injoffice/')) {
    const segments = specifier.split('/')
    const packageName = segments.slice(0, 2).join('/')
    const subpath = segments.slice(2).join('/') || 'index'
    for (const manifestPath of [...tree.keys()].filter((path) => /^packages\/[^/]+\/package\.json$/.test(path)).sort()) {
      const packageManifest = JSON.parse(readBaselineFile(root, manifestPath, `baseline package manifest ${manifestPath}`, []))
      if (packageManifest.name !== packageName) continue
      const entry = packageManifest.injoffice?.entries?.[subpath] ?? (subpath === 'index' ? 'src/index.ts' : `src/${subpath}.ts`)
      candidates.push(posix.normalize(posix.join(posix.dirname(manifestPath), entry)))
    }
  }
  for (const candidate of candidates) {
    const sourcePath = baselineSourcePath(tree, candidate)
    if (sourcePath) return sourcePath
  }
  return undefined
}

function statDirectoryNames(path) {
  return existsSync(path)
    ? readFileDirectory(path).filter((entry) => entry.isDirectory()).map((entry) => entry.name).sort()
    : []
}

function readFileDirectory(path) {
  // Kept behind a tiny wrapper so validation remains read-only and easy to stub.
  return readdirSync(path, { withFileTypes: true })
}

function scanAuthorityClosure(root, initialPaths, errors, { baseline = false } = {}) {
  const queue = [...initialPaths]
  const visited = new Set()
  let tree
  if (baseline) {
    try {
      tree = readBaselineTree(root)
    } catch (error) {
      errors.push(error instanceof Error ? error.message : String(error))
      return
    }
  }
  while (queue.length > 0) {
    const repositoryPath = queue.shift()
    if (visited.has(repositoryPath)) continue
    visited.add(repositoryPath)
    const path = baseline ? baselineSourcePath(tree, repositoryPath) : safePath(root, repositoryPath, 'native authority source', errors)
    if (!path) {
      if (baseline) errors.push(`native authority source is absent or not a regular file in baseline ${activeBaseline}: ${repositoryPath}`)
      continue
    }
    if (!/\.(?:[cm]?js|jsx|[cm]?ts|tsx|go)$/.test(path)) continue
    const source = baseline
      ? readBaselineFile(root, repositoryPath, `native authority source ${repositoryPath}`, errors)?.toString('utf8')
      : readFileSync(path, 'utf8')
    if (source === undefined) continue
    const sourceWithoutComments = normalizeSourceComments(source)
    for (const hit of findForbiddenNativeAuthority(sourceWithoutComments)) errors.push(`native authority ${repositoryPath} uses forbidden ${hit} authority`)
    if (!/\.(?:[cm]?js|jsx|[cm]?ts|tsx)$/.test(path)) continue
    for (const specifier of staticSourceDependencies(sourceWithoutComments)) {
      if (baseline) {
        const imported = resolveBaselineSourceImport(root, tree, repositoryPath, specifier)
        if (imported && !visited.has(imported)) queue.push(imported)
        continue
      }
      const imported = resolveSourceImport(root, path, specifier)
      if (!imported) continue
      const realRoot = realpathSync(root)
      const realImported = realpathSync(imported)
      if (!containedPath(realRoot, realImported)) {
        errors.push(`native authority import ${specifier} from ${repositoryPath} resolves outside the repository`)
        continue
      }
      const repositoryImport = relative(realRoot, realImported).split(sep).join('/')
      if (!visited.has(repositoryImport)) queue.push(repositoryImport)
    }
  }
}

export function validateNativeAuthorityClosure(initialPaths, { root = scriptRoot } = {}) {
  const errors = []
  scanAuthorityClosure(root, initialPaths, errors)
  return [...new Set(errors)].sort()
}

export function validateCompletionManifest(manifest, { root = scriptRoot } = {}) {
  const errors = []
  const fixtureByID = new Map()
  const referencedFixtures = new Set()
  const authorityPaths = new Set()
  const corpusByPath = new Map()

  const schemaErrors = validatePublishedSchema(root, manifest)
  if (schemaErrors.length > 0) return [...new Set(schemaErrors)].sort()

  const corpusManifestPath = safePath(root, 'go/officecompat/corpus/manifest.json', 'canonical corpus manifest', errors)
  if (corpusManifestPath) {
    try {
      const corpusBytes = readBaselineFile(root, 'go/officecompat/corpus/manifest.json', 'canonical corpus manifest', errors)
      const corpusManifest = JSON.parse(corpusBytes?.toString('utf8') ?? '')
      for (const fixture of corpusManifest.fixtures ?? []) corpusByPath.set(`go/officecompat/corpus/${fixture.package}`, fixture)
    } catch (error) {
      errors.push(`canonical corpus manifest is invalid JSON: ${error instanceof Error ? error.message : String(error)}`)
    }
  }

  if (!exactKeys(manifest, ['protocol', 'version', 'baseline', 'policies', 'fixtures', 'qualifications', 'capabilities', 'pendingWork'], 'manifest', errors)) return errors
  const expectedProtocol = protocolByVersion[manifest.version]
  const expectedBaseline = baselineByVersion[manifest.version]
  if (!expectedProtocol || !expectedBaseline) errors.push(`manifest version must be 1, 2, or 3`)
  else if (manifest.protocol !== expectedProtocol) errors.push(`manifest protocol must be ${expectedProtocol}`)
  if (expectedBaseline) activeBaseline = expectedBaseline

  if (exactKeys(manifest.baseline, ['repository', 'ref', 'commit'], 'baseline', errors)) {
    if (manifest.baseline.repository !== 'injectinglabs/injoffice' || manifest.baseline.ref !== 'origin/main' || manifest.baseline.commit !== expectedBaseline) {
      errors.push(`baseline must remain exact origin/main ${expectedBaseline}; update requires a new matrix version`)
    }
  }

  if (exactKeys(manifest.policies, ['completionRequirements', 'requiredDimensions', 'nativeProductionEntrypoints'], 'policies', errors)) {
    if (exactKeys(manifest.policies.completionRequirements, ['officeExportFixture', 'productionE2E'], 'policies.completionRequirements', errors)) {
      for (const requirement of ['officeExportFixture', 'productionE2E']) if (manifest.policies.completionRequirements[requirement] !== true) errors.push(`policies.completionRequirements.${requirement} must remain true`)
    }
    if (!Array.isArray(manifest.policies.requiredDimensions) || !DIMENSIONS.every((dimension) => manifest.policies.requiredDimensions.includes(dimension)) || !sortedUnique(manifest.policies.requiredDimensions)) {
      errors.push(`policies.requiredDimensions must be sorted, unique, and include ${DIMENSIONS.join(', ')}`)
    }
    if (!sortedUnique(manifest.policies.nativeProductionEntrypoints) || manifest.policies.nativeProductionEntrypoints.length === 0) errors.push('policies.nativeProductionEntrypoints must be a non-empty sorted unique list')
    for (const entrypoint of manifest.policies.nativeProductionEntrypoints ?? []) {
      authorityPaths.add(entrypoint)
      if (safePath(root, entrypoint, 'native production entrypoint', errors)) readBaselineFile(root, entrypoint, 'native production entrypoint', errors)
    }
  }

  if (!Array.isArray(manifest.fixtures) || manifest.fixtures.length === 0) errors.push('fixtures must be a non-empty array')
  let previousFixture = ''
  for (const [index, fixture] of (manifest.fixtures ?? []).entries()) {
    const context = `fixtures[${index}]`
    if (!exactKeys(fixture, ['id', 'format', 'kind', 'path', 'sha256', 'provenance'], context, errors)) continue
    if (!idPattern.test(fixture.id ?? '') || fixture.id <= previousFixture) errors.push(`${context}.id must be canonical and strictly ordered`)
    previousFixture = fixture.id
    if (fixtureByID.has(fixture.id)) errors.push(`${context}.id is duplicated`)
    fixtureByID.set(fixture.id, fixture)
    if (![...FORMATS, 'shared'].includes(fixture.format)) errors.push(`${context}.format is invalid`)
    if (!['generated', 'generated-contract', 'generated-inline', 'office-export'].includes(fixture.kind)) errors.push(`${context}.kind is invalid`)
    const path = safePath(root, fixture.path, `${context}.path`, errors)
    if (!shaPattern.test(fixture.sha256 ?? '')) errors.push(`${context}.sha256 must be canonical lowercase SHA-256`)
    if (path && shaPattern.test(fixture.sha256 ?? '')) {
      const baselineBytes = readBaselineFile(root, fixture.path, context, errors)
      const actual = baselineBytes ? canonicalSha(baselineBytes) : undefined
      if (actual !== fixture.sha256) errors.push(`${context}.sha256 is stale for ${fixture.path}: expected ${actual}`)
    }
    if (exactKeys(fixture.provenance, ['producer', 'source', 'license'], `${context}.provenance`, errors)) {
      for (const key of ['producer', 'source', 'license']) if (typeof fixture.provenance[key] !== 'string' || fixture.provenance[key].length === 0) errors.push(`${context}.provenance.${key} is required`)
    }
    if (fixture.kind === 'generated') {
      const corpusFixture = corpusByPath.get(fixture.path)
      if (!corpusFixture) errors.push(`${context} generated fixture is not owned by the canonical officecompat corpus`)
      else {
        if (fixture.id !== corpusFixture.id || fixture.format !== corpusFixture.format || fixture.sha256 !== corpusFixture.sha256) errors.push(`${context} drifts from canonical corpus identity or digest`)
        if (fixture.provenance?.producer !== corpusFixture.provenance?.generator || fixture.provenance?.source !== corpusFixture.provenance?.source || fixture.provenance?.license !== corpusFixture.provenance?.license) errors.push(`${context} drifts from canonical corpus provenance`)
      }
    }
  }

  if (exactKeys(manifest.qualifications, ['differential', 'resource'], 'qualifications', errors)) {
    for (const qualificationName of ['differential', 'resource']) {
      const qualification = manifest.qualifications[qualificationName]
      const context = `qualifications.${qualificationName}`
      if (!exactKeys(qualification, ['status', 'owner', 'evidence', 'limitations'], context, errors)) continue
      if (qualification.status !== 'present') errors.push(`${context}.status must be present for this baseline`)
      if (typeof qualification.owner !== 'string' || qualification.owner.length === 0) errors.push(`${context}.owner is required`)
      requireStrings(qualification.limitations, `${context}.limitations`, errors)
      if (!Array.isArray(qualification.evidence) || qualification.evidence.length === 0) errors.push(`${context}.evidence must be non-empty`)
      for (const [evidenceIndex, evidence] of (qualification.evidence ?? []).entries()) {
        const evidenceContext = `${context}.evidence[${evidenceIndex}]`
        if (!exactKeys(evidence, ['path', 'selector', 'fixtures'], evidenceContext, errors)) continue
        const path = safePath(root, evidence.path, `${evidenceContext}.path`, errors)
        const baselineEvidence = path ? readBaselineFile(root, evidence.path, evidenceContext, errors)?.toString('utf8') : undefined
        if (typeof evidence.selector !== 'string' || evidence.selector.length < 12) errors.push(`${evidenceContext}.selector must be a specific baseline anchor`)
        else if (baselineEvidence !== undefined && !baselineEvidence.includes(evidence.selector)) errors.push(`${evidenceContext}.selector is stale in baseline ${evidence.path}`)
        if (typeof evidence.path === 'string' && evidence.path.endsWith('_test.go') && typeof evidence.selector === 'string' && !/^(?:Test|Fuzz)[A-Za-z0-9_]+$/.test(evidence.selector)) errors.push(`${evidenceContext}.selector must name an exact Go test or fuzz target`)
        if (!Array.isArray(evidence.fixtures) || evidence.fixtures.length === 0 || new Set(evidence.fixtures).size !== evidence.fixtures.length) errors.push(`${evidenceContext}.fixtures must be non-empty and unique`)
        for (const fixtureID of evidence.fixtures ?? []) {
          const fixture = fixtureByID.get(fixtureID)
          if (!fixture) errors.push(`${evidenceContext} references unknown fixture ${fixtureID}`)
          else {
            referencedFixtures.add(fixtureID)
            if (fixture.kind === 'generated-inline' && fixture.path !== evidence.path) errors.push(`${evidenceContext} does not execute inline generator fixture ${fixtureID}`)
            if (fixture.kind === 'generated-contract' && !(baselineEvidence ?? '').includes(basename(fixture.path))) errors.push(`${evidenceContext} does not reference contract fixture ${fixtureID}`)
          }
        }
      }
    }
  }

  if (!Array.isArray(manifest.capabilities) || manifest.capabilities.length === 0) errors.push('capabilities must be a non-empty array')
  let previousCapability = ''
  const dimensionsByFormat = new Map(FORMATS.map((format) => [format, new Set()]))
  for (const [index, capability] of (manifest.capabilities ?? []).entries()) {
    const context = `capabilities[${index}]`
    const capabilityFixtures = new Set()
    if (!exactKeys(capability, ['id', 'format', 'status', 'summary', 'dimensions', 'authoritativeModules', 'semantics', 'gates'], context, errors)) continue
    if (!capabilityPattern.test(capability.id ?? '') || capability.id <= previousCapability) errors.push(`${context}.id must be canonical and strictly ordered`)
    previousCapability = capability.id
    if (!FORMATS.includes(capability.format) || !capability.id.startsWith(`${capability.format}.`)) errors.push(`${context} format/id mismatch`)
    if (!['complete', 'partial', 'pending'].includes(capability.status)) errors.push(`${context}.status is invalid`)
    if (typeof capability.summary !== 'string' || capability.summary.length === 0) errors.push(`${context}.summary is required`)
    if (requireStrings(capability.dimensions, `${context}.dimensions`, errors, { sorted: true })) for (const dimension of capability.dimensions) {
      if (!CAPABILITY_DIMENSIONS.includes(dimension)) errors.push(`${context}.dimensions contains unknown dimension ${dimension}`)
      dimensionsByFormat.get(capability.format)?.add(dimension)
    }
    if (requireStrings(capability.authoritativeModules, `${context}.authoritativeModules`, errors, { sorted: true })) for (const modulePath of capability.authoritativeModules) {
      authorityPaths.add(modulePath)
      if (safePath(root, modulePath, `${context}.authoritativeModules`, errors)) readBaselineFile(root, modulePath, `${context}.authoritativeModules`, errors)
    }

    if (exactKeys(capability.semantics, ['supported', 'preserved', 'refused', 'silentLossPolicy'], `${context}.semantics`, errors)) {
      for (const field of ['supported', 'preserved', 'refused']) {
        requireStrings(capability.semantics[field], `${context}.semantics.${field}`, errors)
        if (capability.semantics[field]?.length !== 1) errors.push(`${context}.semantics.${field} must be one atomic promise in v1`)
      }
      if (!['preserve-verbatim', 'refuse-atomically', 'preserve-or-refuse'].includes(capability.semantics.silentLossPolicy)) errors.push(`${context}.semantics.silentLossPolicy is invalid`)
    }

    if (!exactKeys(capability.gates, GATES, `${context}.gates`, errors)) continue
    for (const gateName of GATES) {
      const gate = capability.gates[gateName]
      const gateContext = `${context}.gates.${gateName}`
      if (!exactKeys(gate, ['status', 'owner', 'evidence'], gateContext, errors, ['reason'])) continue
      if (!['present', 'missing', 'not-applicable'].includes(gate.status)) errors.push(`${gateContext}.status is invalid`)
      if (typeof gate.owner !== 'string' || gate.owner.length === 0) errors.push(`${gateContext}.owner is required`)
      if (!Array.isArray(gate.evidence)) errors.push(`${gateContext}.evidence must be an array`)
      if (gate.status === 'present' && (!Array.isArray(gate.evidence) || gate.evidence.length === 0)) errors.push(`${gateContext} is present without evidence`)
      if (gate.status !== 'present' && Array.isArray(gate.evidence) && gate.evidence.length > 0) errors.push(`${gateContext} has evidence but is ${gate.status}`)
      if (gate.status !== 'present' && (typeof gate.reason !== 'string' || gate.reason.length === 0)) errors.push(`${gateContext} ${gate.status} requires a reason`)
      if (gate.status === 'present' && 'reason' in gate) errors.push(`${gateContext} present gate must not carry a reason`)
      for (const [evidenceIndex, evidence] of (gate.evidence ?? []).entries()) {
        const evidenceContext = `${gateContext}.evidence[${evidenceIndex}]`
        if (!exactKeys(evidence, ['path', 'selector', 'fixtures'], evidenceContext, errors)) continue
        const path = safePath(root, evidence.path, `${evidenceContext}.path`, errors)
        const baselineEvidence = path ? readBaselineFile(root, evidence.path, evidenceContext, errors)?.toString('utf8') : undefined
        if (typeof evidence.selector !== 'string' || evidence.selector.length < 3) errors.push(`${evidenceContext}.selector is required`)
        else if (baselineEvidence !== undefined && !baselineEvidence.includes(evidence.selector)) errors.push(`${evidenceContext}.selector is stale in baseline ${evidence.path}`)
        if (typeof evidence.selector === 'string' && evidence.selector.length < 12) errors.push(`${evidenceContext}.selector is too generic to bind a test`)
        if (typeof evidence.path === 'string' && evidence.path.endsWith('_test.go') && typeof evidence.selector === 'string' && !/^(?:Test|Fuzz)[A-Za-z0-9_]+$/.test(evidence.selector)) errors.push(`${evidenceContext}.selector must name an exact Go test or fuzz target`)
        if (gateName === 'productionE2E' && gate.status === 'present' && typeof evidence.path === 'string' && !PRODUCTION_E2E_EVIDENCE_PATHS.includes(evidence.path)) {
          errors.push(`${evidenceContext}.path must bind the host-owned production E2E kit`)
        }
        if (!Array.isArray(evidence.fixtures) || evidence.fixtures.length === 0 || new Set(evidence.fixtures).size !== evidence.fixtures.length) errors.push(`${evidenceContext}.fixtures must be non-empty and unique`)
        const evidenceSource = baselineEvidence ?? ''
        for (const fixtureID of evidence.fixtures ?? []) {
          const fixture = fixtureByID.get(fixtureID)
          if (!fixture) errors.push(`${evidenceContext} references unknown fixture ${fixtureID}`)
          else if (fixture.format !== 'shared' && fixture.format !== capability.format) errors.push(`${evidenceContext} uses ${fixture.format} fixture ${fixtureID} for ${capability.format}`)
          else {
            referencedFixtures.add(fixtureID)
            capabilityFixtures.add(fixtureID)
            if (fixture.kind === 'generated-inline' && fixture.path !== evidence.path) errors.push(`${evidenceContext} does not execute inline generator fixture ${fixtureID}`)
            if (fixture.kind === 'generated-contract' && !evidenceSource.includes(basename(fixture.path))) errors.push(`${evidenceContext} does not reference contract fixture ${fixtureID}`)
            if (fixture.kind === 'generated' && evidence.path !== `go/officecompat/${capability.format}_corpus_test.go`) errors.push(`${evidenceContext} does not consume canonical corpus fixture ${fixtureID}`)
          }
        }
      }
    }

    if (capability.gates.structuralPreservation?.status !== 'present') errors.push(`${context} has no structural preservation test`)
    if (capability.gates.unsupportedObjectLoss?.status !== 'present') errors.push(`${context} has no silent unsupported-object loss gate`)
    // v1 intentionally permits one atomic promise in each semantic class. These
    // fixed bindings make every promise test-addressable without free-form refs:
    // supported/preserved -> structuralPreservation; refused -> unsupportedObjectLoss.
    if (capability.semantics?.supported?.length === 1 && capability.gates.structuralPreservation?.status !== 'present') errors.push(`${context}.semantics.supported is not bound to structural evidence`)
    if (capability.semantics?.preserved?.length === 1 && capability.gates.structuralPreservation?.status !== 'present') errors.push(`${context}.semantics.preserved is not bound to structural evidence`)
    if (capability.semantics?.refused?.length === 1 && capability.gates.unsupportedObjectLoss?.status !== 'present') errors.push(`${context}.semantics.refused is not bound to refusal evidence`)
    if (capability.dimensions?.some((value) => value === 'layout' || value === 'visual-production') && capability.gates.renderLayout?.status !== 'present') errors.push(`${context} promises layout/visual production without a render/layout test`)
    if (capability.dimensions?.includes('mutation') && capability.gates.mutationRoundtrip?.status !== 'present') errors.push(`${context} promises mutation without a roundtrip test`)
    if (capability.dimensions?.includes('package-consumer') && capability.gates.packageConsumer?.status !== 'present') errors.push(`${context} promises a package consumer without a consumer test`)
    if (capability.dimensions?.includes('resource') && manifest.qualifications?.resource?.status !== 'present') errors.push(`${context} promises resource qualification without the baseline resource gate`)
    if (capability.status === 'complete') {
      for (const gateName of GATES) if (capability.gates[gateName]?.status === 'missing') errors.push(`${context} is complete before ${gateName} exists`)
      if (manifest.policies.completionRequirements?.productionE2E && capability.gates.productionE2E?.status !== 'present') errors.push(`${context} is complete before production E2E exists`)
      if (manifest.policies.completionRequirements?.officeExportFixture && ![...capabilityFixtures].some((fixtureID) => fixtureByID.get(fixtureID)?.kind === 'office-export')) errors.push(`${context} is complete without evidence from a real Office-export fixture`)
    }
  }

  scanAuthorityClosure(root, [...authorityPaths].sort(), errors, { baseline: true })

  for (const format of FORMATS) for (const dimension of manifest.policies?.requiredDimensions ?? []) {
    if (!dimensionsByFormat.get(format)?.has(dimension)) errors.push(`${format} has no capability for required ${dimension} dimension`)
  }
  for (const fixtureID of fixtureByID.keys()) if (!referencedFixtures.has(fixtureID)) errors.push(`fixture ${fixtureID} is not referenced by matrix evidence`)

  if (!Array.isArray(manifest.pendingWork)) errors.push('pendingWork must be an offline snapshot for this matrix version')
  let previousPending = ''
  const pendingPRs = new Set()
  for (const [index, pending] of (manifest.pendingWork ?? []).entries()) {
    const context = `pendingWork[${index}]`
    if (!exactKeys(pending, ['id', 'format', 'status', 'pr', 'base', 'branch', 'head', 'targets'], context, errors, ['note'])) continue
    if (!idPattern.test(pending.id ?? '') || pending.id <= previousPending) errors.push(`${context}.id must be canonical and strictly ordered`)
    previousPending = pending.id
    if (![...FORMATS, 'shared'].includes(pending.format)) errors.push(`${context}.format is invalid`)
    if (pending.status !== 'pending') errors.push(`${context}.status must remain pending until merged into a later baseline`)
    if (!Number.isSafeInteger(pending.pr) || pending.pr <= 0 || pendingPRs.has(pending.pr)) errors.push(`${context}.pr must be a unique positive integer`)
    pendingPRs.add(pending.pr)
    if (typeof pending.base !== 'string' || !/^main@[a-f0-9]{40}$/.test(pending.base)) errors.push(`${context}.base must be main at an exact 40-character Git SHA`)
    if (typeof pending.branch !== 'string' || !/^(?:feat|fix|test|docs|chore)\/[a-z0-9]+(?:[./-][a-z0-9]+)*$/.test(pending.branch)) errors.push(`${context}.branch must be canonical`)
    if (typeof pending.head === 'string' && !/^[a-f0-9]{40}$/.test(pending.head)) errors.push(`${context}.head must be an exact 40-character Git SHA`)
    requireStrings(pending.targets, `${context}.targets`, errors, { sorted: true })
    if (typeof pending.branch === 'string' && typeof pending.head === 'string' && /^[a-f0-9]{40}$/.test(pending.head)) {
      const refs = [`refs/remotes/origin/${pending.branch}`, `refs/heads/${pending.branch}`]
      let localHead
      for (const ref of refs) {
        const result = spawnSync('git', ['-C', root, 'rev-parse', '--verify', `${ref}^{commit}`], { encoding: 'utf8' })
        if (result.status === 0) { localHead = result.stdout.trim(); break }
      }
      if (!localHead) errors.push(`${context}.branch has no exact local ref; fetch history before validating`)
      else if (localHead !== pending.head) errors.push(`${context}.head is stale for local ${pending.branch}: expected ${localHead}`)
      else if (/^main@[a-f0-9]{40}$/.test(pending.base ?? '')) {
        const result = spawnSync('git', ['-C', root, 'merge-base', BASELINE, pending.head], { encoding: 'utf8' })
        const actualBase = result.status === 0 ? result.stdout.trim() : undefined
        if (actualBase !== pending.base.slice(5)) errors.push(`${context}.base is stale for local ${pending.branch}: expected main@${actualBase ?? 'unavailable'}`)
      }
    }
  }

  return [...new Set(errors)].sort()
}

export function loadCompletionManifest(path = resolve(scriptRoot, 'testdata/native-office-completion/v1/manifest.json')) {
  return JSON.parse(readFileSync(path, 'utf8'))
}

export function validateCanonicalCompletionManifest({ root = scriptRoot } = {}) {
  const versions = [
    ['v1', 'testdata/native-office-completion/v1/manifest.json'],
    ['v2', 'testdata/native-office-completion/v2/manifest.json'],
    ['v3', 'testdata/native-office-completion/v3/manifest.json'],
  ]
  const errors = []
  for (const [label, relativePath] of versions) {
    for (const error of validateCompletionManifest(loadCompletionManifest(resolve(root, relativePath)), { root })) {
      errors.push(`${label}: ${error}`)
    }
  }
  return errors
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  if (process.argv.includes('--print-fixture-digests')) {
    const manifest = loadCompletionManifest()
    for (const fixture of manifest.fixtures) {
      const errors = []
      const bytes = readBaselineFile(scriptRoot, fixture.path, `fixture ${fixture.id}`, errors)
      if (!bytes) throw new Error(errors.join('; '))
      console.log(`${fixture.id}\t${canonicalSha(bytes)}\t${fixture.path}`)
    }
    process.exit(0)
  }
  const errors = validateCanonicalCompletionManifest()
  if (errors.length > 0) {
    console.error(errors.map((error) => `- ${error}`).join('\n'))
    process.exitCode = 1
  } else {
    for (const version of [1, 2, 3]) {
      const manifest = loadCompletionManifest(resolve(scriptRoot, `testdata/native-office-completion/v${version}/manifest.json`))
      console.log(`Validated native Office completion matrix v${manifest.version}: ${manifest.capabilities.length} capabilities, ${manifest.fixtures.length} provenance-bound fixtures, ${manifest.pendingWork.length} pending lanes.`)
    }
  }
}
