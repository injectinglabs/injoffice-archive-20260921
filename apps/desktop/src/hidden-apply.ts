/** Debounced native apply. Skip while IME composition is active. Never treat the preview as file authority. */
export function createHiddenApplyScheduler({
  delayMs = 80,
  composing,
  apply,
}: {
  delayMs?: number;
  composing(): boolean;
  apply(): Promise<void> | void;
}) {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let running: Promise<void> | undefined;
  let pending = false;

  const run = () => {
    timer = undefined;
    if (composing()) { pending = true; return; }
    pending = false;
    running = Promise.resolve(apply()).finally(() => {
      running = undefined;
      if (pending && !composing()) queue();
    });
    return running;
  };

  const queue = () => {
    pending = true;
    if (running || timer || composing()) return;
    timer = setTimeout(run, delayMs);
  };

  return {
    schedule: queue,
    async flush() {
      if (timer) { clearTimeout(timer); timer = undefined; }
      if (running) await running;
      if (pending && !composing()) await run();
    },
    cancel() {
      pending = false;
      if (timer) { clearTimeout(timer); timer = undefined; }
    },
  };
}
