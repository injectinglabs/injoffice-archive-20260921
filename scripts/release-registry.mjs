// Version metadata endpoints require JSON, not the abbreviated install packument.
export async function publishedIntegrity(name, version, fetcher = fetch) {
  const response = await fetcher(`https://registry.npmjs.org/${encodeURIComponent(name)}/${version}`, {
    headers: { accept: 'application/json' },
    signal: AbortSignal.timeout(30_000),
  })
  if (response.status === 404) return null
  if (!response.ok) throw new Error(`${name}@${version}: registry preflight failed with HTTP ${response.status}`)
  const metadata = await response.json()
  if (typeof metadata.dist?.integrity !== 'string' || !metadata.dist.integrity) {
    throw new Error(`${name}@${version}: registry response is missing dist.integrity`)
  }
  return metadata.dist.integrity
}
