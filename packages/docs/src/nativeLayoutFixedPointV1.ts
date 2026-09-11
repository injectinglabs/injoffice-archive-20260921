/** Shared bounded solve for pagination-dependent field text and line exclusions. */
export async function solveNativeDocxLayoutFixedPointV1<State, Result>(
  seed: State,
  step: (state: State, pass: number) => Promise<{ next: State; result: Result }>,
): Promise<{ state: State; result: Result; passes: number }> {
  const key = (value: State): string => {
    const canonical = (value: unknown): unknown => {
      if (Array.isArray(value)) return value.map(canonical)
      if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).sort(([a],[b]) => a < b ? -1 : a > b ? 1 : 0).map(([k,v]) => [k,canonical(v)]))
      if (value === undefined || typeof value === 'number' && !Number.isFinite(value)) throw new TypeError('Layout fixed-point state must be finite JSON')
      return value
    }
    const serialized = JSON.stringify(canonical(value))
    if (serialized.length > 1_000_000) throw new RangeError('Layout fixed-point state exceeds byte budget')
    return serialized
  }
  let state = structuredClone(seed), signature = key(state)
  const seen = new Set([signature])
  for (let pass = 0; pass < 8; pass++) {
    const { next, result } = await step(structuredClone(state), pass)
    const following = key(next)
    if (following === signature) return { state: structuredClone(next), result, passes: pass + 1 }
    if (seen.has(following)) throw new TypeError('Pagination-dependent layout entered a cycle; native preview refused')
    seen.add(following); state = structuredClone(next); signature = following
  }
  throw new TypeError('Pagination-dependent layout did not converge within eight passes; native preview refused')
}
