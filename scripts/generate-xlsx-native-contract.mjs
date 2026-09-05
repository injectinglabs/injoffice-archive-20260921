import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import { readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(fileURLToPath(new URL('.', import.meta.url)), '..')
const check = process.argv.includes('--check')
const contracts = [
  {
    version: 1,
    schemaPath: resolve(root, 'schemas/xlsx-native-v1.schema.json'),
    tsPath: resolve(root, 'packages/sheets/src/nativeContract.generated.ts'),
    goPath: resolve(root, 'go/xlsxpatch/native_contract_schema_generated.go'),
    sourceLabel: 'schemas/xlsx-native-v1.schema.json',
    tsPrefix: 'XLSX_NATIVE',
    goExportPrefix: 'NativeXLSXSchema',
    goInternalPrefix: 'nativeXLSXSchema',
    goBindings: 'nativeXLSXBindingShapes',
    goClassifications: 'nativeXLSXSchemaUnsupportedClassifications',
  },
  {
    version: 2,
    schemaPath: resolve(root, 'schemas/xlsx-native-v2.schema.json'),
    tsPath: resolve(root, 'packages/sheets/src/nativeContractV2.generated.ts'),
    goPath: resolve(root, 'go/xlsxpatch/native_contract_v2_schema_generated.go'),
    sourceLabel: 'schemas/xlsx-native-v2.schema.json',
    tsPrefix: 'XLSX_NATIVE_V2',
    goExportPrefix: 'NativeXLSXV2Schema',
    goInternalPrefix: 'nativeXLSXV2Schema',
    goBindings: 'nativeXLSXV2BindingShapes',
    goClassifications: 'nativeXLSXV2SchemaUnsupportedClassifications',
  },
]

const pascal = (value) => value.split(/[^A-Za-z0-9]+/).filter(Boolean).map((part) => part[0].toUpperCase() + part.slice(1)).join('')
const goImpact = (value) => value === 'none' ? 'No' : pascal(value)

async function buildContract(config) {
  const source = await readFile(config.schemaPath, 'utf8')
  const schema = JSON.parse(source)
  const digest = createHash('sha256').update(source).digest('hex')
  const limits = schema['x-resource-limits']
  const unsupportedClassifications = schema['x-unsupported-classifications']

  if (schema.properties?.version?.const !== config.version) throw new Error(`${config.schemaPath} has the wrong protocol version`)
  if (!limits || Object.values(limits).some((value) => !Number.isSafeInteger(value) || value <= 0)) throw new Error('x-resource-limits must contain positive safe integers')
  if (!unsupportedClassifications || Object.entries(unsupportedClassifications).some(([code, value]) =>
    !/^[A-Z][A-Z0-9_]{0,127}$/.test(code) || typeof value?.capability !== 'string' ||
    !['workbook', 'sheet', 'style'].includes(value.scope) || !['none', 'cell', 'range', 'sheet'].includes(value.impact) ||
    (value.impact === 'sheet' && typeof value.refusalCode !== 'string') || (value.impact !== 'sheet' && value.refusalCode !== undefined))) {
    throw new Error('x-unsupported-classifications is invalid')
  }

  const definitions = new Map([['$root', schema], ...Object.entries(schema.$defs ?? {})])
  const named = [...definitions.entries()].filter(([, value]) => typeof value['x-binding-name'] === 'string')
  const objects = named.filter(([, value]) => value.type === 'object')
  for (const [name, value] of definitions) {
    if (value.type === 'object' && typeof value['x-binding-name'] !== 'string') throw new Error(`object ${name} is missing x-binding-name`)
  }

  function refDefinition(ref) {
    const prefix = '#/$defs/'
    if (!ref.startsWith(prefix)) throw new Error(`unsupported reference ${ref}`)
    const target = definitions.get(ref.slice(prefix.length))
    if (!target?.['x-binding-name']) throw new Error(`reference ${ref} has no binding name`)
    return target
  }
  function wireType(value) {
    if ('$ref' in value) {
      const target = refDefinition(value.$ref)
      return target.type === 'object' ? target['x-binding-name'] : wireType(target)
    }
    if ('const' in value) return Number.isInteger(value.const) ? 'integer' : typeof value.const
    if (value.type === 'array') return `[]${wireType(value.items)}`
    if (value.type === 'integer' || value.type === 'number' || value.type === 'string' || value.type === 'boolean') return value.type
    throw new Error(`unsupported schema wire type ${JSON.stringify(value)}`)
  }
  const objectBindings = Object.fromEntries(objects.map(([schemaName, value]) => [value['x-binding-name'], {
    schemaName,
    properties: Object.keys(value.properties ?? {}).sort(),
    required: [...(value.required ?? [])].sort(),
    types: Object.fromEntries(Object.entries(value.properties ?? {}).sort(([left], [right]) => left.localeCompare(right)).map(([name, child]) => [name, wireType(child)])),
  }]).sort(([left], [right]) => left.localeCompare(right)))

  function tsType(value) {
    if ('$ref' in value) return refDefinition(value.$ref)['x-binding-name']
    if ('const' in value) return JSON.stringify(value.const)
    if (Array.isArray(value.enum)) return value.enum.map(JSON.stringify).join(' | ')
    if (value.type === 'array') return `ReadonlyArray<${tsType(value.items)}>`
    if (value.type === 'string') return 'string'
    if (value.type === 'integer' || value.type === 'number') return 'number'
    if (value.type === 'boolean') return 'boolean'
    throw new Error(`unsupported schema shape ${JSON.stringify(value)}`)
  }
  const typeDeclarations = named.map(([, value]) => {
    const name = value['x-binding-name']
    if (value.type === 'string') return `export type ${name} = ${value.enum.map(JSON.stringify).join(' | ')}`
    const required = new Set(value.required ?? [])
    const fields = Object.entries(value.properties ?? {}).map(([field, child]) => `  readonly ${field}${required.has(field) ? '' : '?'}: ${tsType(child)}`)
    return `export interface ${name} {\n${fields.join('\n')}\n}`
  }).join('\n\n')

  const schemaIdentity = config.version === 1 ? digest : `sha256:${digest}`
  let ts = `// Code generated by scripts/generate-xlsx-native-contract.mjs; DO NOT EDIT.\n// Source: ${config.sourceLabel}\n\n` +
    `export const ${config.tsPrefix}_SCHEMA_ID = ${JSON.stringify(schema.$id)} as const\n` +
    `export const ${config.tsPrefix}_SCHEMA_SHA256 = ${JSON.stringify(schemaIdentity)} as const\n` +
    `export const ${config.tsPrefix}_PROTOCOL = ${JSON.stringify(schema.properties.protocol.const)} as const\n` +
    `export const ${config.tsPrefix}_VERSION = ${JSON.stringify(schema.properties.version.const)} as const\n`
  if (schema['x-media-type']) ts += `export const ${config.tsPrefix}_MEDIA_TYPE = ${JSON.stringify(schema['x-media-type'])} as const\n`
  ts += `export const ${config.tsPrefix}_RESOURCE_LIMITS = ${JSON.stringify(limits, null, 2)} as const\n` +
    `export const ${config.tsPrefix}_UNSUPPORTED_CLASSIFICATIONS = ${JSON.stringify(unsupportedClassifications, null, 2)} as const\n` +
    `export const ${config.tsPrefix}_OBJECT_BINDINGS = ${JSON.stringify(objectBindings, null, 2)} as const\n` +
    `export const ${config.tsPrefix}_SCHEMA = ${JSON.stringify(schema, null, 2)} as const\n\n` + typeDeclarations + '\n'

  const goShape = config.version === 1 ? 'nativeXLSXBindingShape' : 'nativeXLSXV2BindingShape'
  let goSource = `// Code generated by scripts/generate-xlsx-native-contract.mjs; DO NOT EDIT.\n// Source: ../../${config.sourceLabel}\n\npackage xlsxpatch\n\n` +
    `const ${config.goExportPrefix}ID = ${JSON.stringify(schema.$id)}\n` +
    `const ${config.goExportPrefix}SHA256 = ${JSON.stringify(schemaIdentity)}\n` +
    `const ${config.goInternalPrefix}Protocol = ${JSON.stringify(schema.properties.protocol.const)}\n` +
    `const ${config.goInternalPrefix}Version = ${JSON.stringify(schema.properties.version.const)}\n`
  if (schema['x-media-type']) goSource += `const NativeXLSXV2MediaType = ${JSON.stringify(schema['x-media-type'])}\n`
  goSource += Object.entries(limits).map(([name, value]) => `const ${config.goInternalPrefix}${pascal(name)} = ${value}`).join('\n') + '\n\n' +
    `type ${goShape} struct {\n\tProperties []string\n\tRequired []string\n\tTypes map[string]string\n}\n\n` +
    `var ${config.goBindings} = map[string]${goShape}{\n` +
    Object.entries(objectBindings).map(([name, binding]) => `\t${JSON.stringify(name)}: {Properties: []string{${binding.properties.map(JSON.stringify).join(', ')}}, Required: []string{${binding.required.map(JSON.stringify).join(', ')}}, Types: map[string]string{${Object.entries(binding.types).map(([field, type]) => `${JSON.stringify(field)}: ${JSON.stringify(type)}`).join(', ')}}},`).join('\n') +
    '\n}\n\n' +
    `var ${config.goClassifications} = map[string]nativeUnsupportedClassification{\n` +
    Object.entries(unsupportedClassifications).map(([code, value]) => `\t${JSON.stringify(code)}: {capability: ${JSON.stringify(value.capability)}, scope: ${JSON.stringify(value.scope)}, impact: nativeUnsupported${goImpact(value.impact)}Impact${value.refusalCode ? `, refusalCode: ${JSON.stringify(value.refusalCode)}` : ''}},`).join('\n') +
    '\n}\n'
  const go = execFileSync('gofmt', { input: goSource, encoding: 'utf8' })
  await emit(config.tsPath, ts)
  await emit(config.goPath, go)
  return digest
}

async function emit(path, contents) {
  if (!contents.endsWith('\n')) contents += '\n'
  if (check) {
    let existing = ''
    try { existing = await readFile(path, 'utf8') } catch {}
    if (existing !== contents) throw new Error(`${path} is stale; run npm run generate:xlsx-contract`)
  } else {
    await writeFile(path, contents)
  }
}

const digests = []
for (const contract of contracts) digests.push(await buildContract(contract))
console.log(`${check ? 'checked' : 'generated'} XLSX native contracts v1=${digests[0]} v2=${digests[1]}`)
