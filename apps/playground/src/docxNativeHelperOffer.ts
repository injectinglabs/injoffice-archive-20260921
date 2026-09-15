const PACKAGE_DIGEST = /^sha256:[0-9a-f]{64}$/

export async function digestNativeDocxPackage(
  bytes: Uint8Array,
  subtle: Pick<SubtleCrypto, 'digest'> = globalThis.crypto.subtle,
): Promise<string> {
  const digest = await subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)
  return `sha256:${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`
}

/** Prefer the extracted document digest so helper identity follows mutate/undo. */
export function nativeDocxHelperPackageDigest(documentDigest: string | undefined, openedDigest: string): string {
  return documentDigest && PACKAGE_DIGEST.test(documentDigest) ? documentDigest : openedDigest
}

/** Replace helper bytes only together with the digest of those bytes. */
export function nativeDocxAdoptOpenedPackage(bytes: Uint8Array, packageDigest: string): { bytes: Uint8Array; openedDigest: string } {
  if (bytes.byteLength < 1 || !PACKAGE_DIGEST.test(packageDigest)) throw new TypeError('Opened DOCX package identity is incomplete.')
  return { bytes, openedDigest: packageDigest }
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
