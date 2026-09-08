/** Only untouched demos may be released automatically. Editors keep their own
 * state, so any interaction pins the instance until an explicit reset/close. */
export function createDemoRetention(release: () => void, delay = 30_000) {
  let timer: ReturnType<typeof setTimeout> | undefined
  let touched = false
  let visible = true
  let loaded = false
  const cancel = () => { clearTimeout(timer); timer = undefined }
  const schedule = () => {
    cancel()
    if (loaded && !visible && !touched) timer = setTimeout(() => {
      loaded = false
      release()
    }, delay)
  }
  return {
    setVisible(value: boolean) { visible = value; schedule() },
    setLoaded(value: boolean) { loaded = value; schedule() },
    touch() { touched = true; cancel() },
    reset() { touched = false; schedule() },
    dispose: cancel,
  }
}
