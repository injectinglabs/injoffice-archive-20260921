const PACKAGE_DIGEST = /^sha256:[0-9a-f]{64}$/

export async function digestNativeDocxPackage(
  bytes: Uint8Array,
  subtle: Pick<SubtleCrypto, 'digest'> = globalThis.crypto.subtle,
): Promise<string> {
  const digest = await subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)
  return `sha256:${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`
}

/** Native page controls bind to opened package bytes, not a successful extract. */
export function nativeDocxHelperOffer(input: {
  apiBase: string
  bytes: Uint8Array | null
  packageDigest: string
}): { bytes: Uint8Array; packageDigest: string; apiBase: string } | null {
  const apiBase = input.apiBase.trim().replace(/\/$/, '')
  if (!apiBase || !input.bytes || input.bytes.byteLength < 1 || !PACKAGE_DIGEST.test(input.packageDigest)) return null
  return { bytes: input.bytes, packageDigest: input.packageDigest, apiBase }
}
